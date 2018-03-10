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
#endregion NameSpace

namespace ReportingTool
{
    public partial class BestSeller : System.Web.UI.Page
    {
        #region Events 

        public DataTable dtBestSeller = null;

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
            if (rdlBestSeller.SelectedValue == "1")
            {
                InsertBestSellerReport();
                InsertBestSellerReportByLineCode7();
            }
            GenerateBestSellerReport();
            GenerateBestSellerReportByLC7();
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
            string filename = ViewState["FileNameBestSeller"].ToString();
            FileDownload(filename);
        }
        #endregion btnDownload_Click

        #region btnBestSellerLC7_Click
        /// <summary>
        /// btnBestSellerLC7_Click
       /// </summary>
       /// <param name="sender"></param>
       /// <param name="e"></param>
        protected void btnBestSellerLC7_Click(object sender, EventArgs e)
        {
            string filename = ViewState["FileNameBestSellerLC7"].ToString();
            FileDownload(filename);
        }
        #endregion btnBestSellerLC7_Click

        #endregion Events


        #region Methods

        #region InsertBestSellerReport
        /// <summary>
        /// Insert BestSeller Report
        /// </summary>
        private void InsertBestSellerReport()
        {
            GetStockDetails objStock = new GetStockDetails();

            //to delete old data
            objStock.DeleteBestSellerReport(); 

            objStock.FromDate = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            objStock.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            //Menswear
            objStock.DivisionCode = "M%";
            objStock.Location = "BAH";
            objStock.InsertBestSellerReport();
            objStock.Location = "JOR";
            
            objStock.InsertBestSellerReport();
            objStock.Location = "OMAN";
            objStock.InsertBestSellerReport();
            objStock.Location = "UAE";
            
            objStock.InsertBestSellerReport();
            objStock.Location = "QAT";
            objStock.InsertBestSellerReport();

            objStock.Location = "KSA";
            objStock.InsertBestSellerReport();

            //Ladieswear
            objStock.DivisionCode = "L%";
            objStock.Location = "BAH";
            objStock.InsertBestSellerReport();
            objStock.Location = "JOR";

            objStock.InsertBestSellerReport();
            objStock.Location = "OMAN";
            objStock.InsertBestSellerReport();
            objStock.Location = "UAE";

            objStock.InsertBestSellerReport();
            objStock.Location = "QAT";
            objStock.InsertBestSellerReport();

            objStock.Location = "KSA";
            objStock.InsertBestSellerReport();

            //chilrenswear
            objStock.DivisionCode = "C%";
            objStock.Location = "BAH";
            objStock.InsertBestSellerReport();
            objStock.Location = "JOR";

            objStock.InsertBestSellerReport();
            objStock.Location = "OMAN";
            objStock.InsertBestSellerReport();
            objStock.Location = "UAE";

            objStock.InsertBestSellerReport();
            objStock.Location = "QAT";
            objStock.InsertBestSellerReport();

            objStock.Location = "KSA";
            objStock.InsertBestSellerReport();

            //Footwear
            objStock.DivisionCode = "F%";
            objStock.Location = "BAH";
            objStock.InsertBestSellerReport();
            objStock.Location = "JOR";

            objStock.InsertBestSellerReport();
            objStock.Location = "OMAN";
            objStock.InsertBestSellerReport();
            objStock.Location = "UAE";

            objStock.InsertBestSellerReport();
            objStock.Location = "QAT";
            objStock.InsertBestSellerReport();

            objStock.Location = "KSA";
            objStock.InsertBestSellerReport();

            //Essentials
            objStock.DivisionCode = "S%";
            objStock.Location = "BAH";
            objStock.InsertBestSellerReport();
            objStock.Location = "JOR";

            objStock.InsertBestSellerReport();
            objStock.Location = "OMAN";
            objStock.InsertBestSellerReport();
            objStock.Location = "UAE";

            objStock.InsertBestSellerReport();
            objStock.Location = "QAT";
            objStock.InsertBestSellerReport();

            objStock.Location = "KSA";
            objStock.InsertBestSellerReport();


            //Homewear
            objStock.DivisionCode = "H%";
            objStock.Location = "BAH";
            objStock.InsertBestSellerReport();
            objStock.Location = "JOR";

            objStock.InsertBestSellerReport();
            objStock.Location = "OMAN";
            objStock.InsertBestSellerReport();
            objStock.Location = "UAE";

            objStock.InsertBestSellerReport();
            objStock.Location = "QAT";
            objStock.InsertBestSellerReport();

            objStock.Location = "KSA";
            objStock.InsertBestSellerReport();

            //Summery

            objStock.DivisionCode = "M%";
            

            objStock.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());
            objStock.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());
            objStock.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
            objStock.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());

            objStock.InsertBestSellerSummeryReport();

            objStock.DivisionCode = "S%";
            objStock.InsertBestSellerSummeryReport();
           
            objStock.DivisionCode = "F%";
            objStock.InsertBestSellerSummeryReport();
            
            objStock.DivisionCode = "C%";
            objStock.InsertBestSellerSummeryReport();
            
            objStock.DivisionCode = "L%";
            objStock.InsertBestSellerSummeryReport();

            objStock.DivisionCode = "H%";
            objStock.InsertBestSellerSummeryReport();

        }
        #endregion InsertBestSellerReport


        #region InsertBestSellerReportByLineCode7
        /// <summary>
        /// Insert Best Seller Report By LineCode7
        /// </summary>
        private void InsertBestSellerReportByLineCode7()
        {
            GetStockDetails objStock = new GetStockDetails();

           
            objStock.FromDate = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            objStock.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            //Menswear
            objStock.DivisionCode = "M%";
            objStock.Location = "BAH";
            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "JOR";

            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "OMAN";
            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "UAE";

            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "QAT";
            objStock.InsertBestSellerByLinecode7();

            objStock.Location = "KSA";
            objStock.InsertBestSellerByLinecode7();

            //Ladieswear
            objStock.DivisionCode = "L%";
            objStock.Location = "BAH";
            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "JOR";

            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "OMAN";
            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "UAE";

            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "QAT";
            objStock.InsertBestSellerByLinecode7();

            objStock.Location = "KSA";
            objStock.InsertBestSellerByLinecode7();

            //chilrenswear
            objStock.DivisionCode = "C%";
            objStock.Location = "BAH";
            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "JOR";

            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "OMAN";
            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "UAE";

            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "QAT";
            objStock.InsertBestSellerByLinecode7();

            objStock.Location = "KSA";
            objStock.InsertBestSellerByLinecode7();

            //Footwear
            objStock.DivisionCode = "F%";
            objStock.Location = "BAH";
            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "JOR";

            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "OMAN";
            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "UAE";

            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "QAT";
            objStock.InsertBestSellerByLinecode7();

            objStock.Location = "KSA";
            objStock.InsertBestSellerByLinecode7();

            //Essentials
            objStock.DivisionCode = "S%";
            objStock.Location = "BAH";
            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "JOR";

            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "OMAN";
            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "UAE";

            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "QAT";
            objStock.InsertBestSellerByLinecode7();

            objStock.Location = "KSA";
            objStock.InsertBestSellerByLinecode7();


            //Homewear
            objStock.DivisionCode = "H%";
            objStock.Location = "BAH";
            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "JOR";

            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "OMAN";
            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "UAE";

            objStock.InsertBestSellerByLinecode7();
            objStock.Location = "QAT";
            objStock.InsertBestSellerByLinecode7();

            objStock.Location = "KSA";
            objStock.InsertBestSellerByLinecode7();


            //Summery

            objStock.DivisionCode = "M%";


            objStock.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());
            objStock.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());
            objStock.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
            objStock.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());

            objStock.InsertBestSellerSummeryReportLC7();

            objStock.DivisionCode = "S%";
            objStock.InsertBestSellerSummeryReportLC7();

            objStock.DivisionCode = "F%";
            objStock.InsertBestSellerSummeryReportLC7();

            objStock.DivisionCode = "C%";
            objStock.InsertBestSellerSummeryReportLC7();

            objStock.DivisionCode = "L%";
            objStock.InsertBestSellerSummeryReportLC7();

            objStock.DivisionCode = "H%";
            objStock.InsertBestSellerSummeryReportLC7();

        }
        #endregion InsertBestSellerReportByLineCode7

        #region GenerateBestSellerReport
        /// <summary>
        /// To generate excel report for Best Seller Report
        /// </summary>
        private void GenerateBestSellerReport()
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




            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\BestSellerReportTemplate2.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);



            //myExcelWorkbook = myExcelApp.Workbooks.Add(1);
            //Excel.Sheets xlSheets1 = myExcelWorkbook.Sheets as Excel.Sheets;

            //Excel.Worksheet xlSheet = (Excel.Worksheet)xlSheets1.Add(xlSheets1[1], Type.Missing, Type.Missing, Type.Missing);



            GetStockDetails objBestSeller = new GetStockDetails();

            //String cellFormulaAsString = myExcelWorksheet.get_Range("A2", misValue).Formula.ToString();
            if (txtFromDate.Text.Trim().Length > 0 && txtToDate.Text.Trim().Length > 0)
            {
                // string WeekNo = txtWeekNo.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\BestSellerReportTemplate2.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];


                objBestSeller.FromDate = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                objBestSeller.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                objBestSeller.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());
                objBestSeller.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());
                objBestSeller.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
                objBestSeller.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());

                objBestSeller.KsaRate = Convert.ToDecimal(txtKsaRate.Text.Trim());

                string ReportType = ddlDivisionCode.SelectedItem.Value;
                if (ReportType == "All")
                {
                    
                    // For Menswear

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "BAH";
                    objBestSeller.DivisionCode = "M%";
                   
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    Excel1.Worksheet xlSheetM = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                    xlSheetM.Name = "Menswear";
                    xlSheetM.get_Range("A2", misValue).Formula = "MENSWEAR BEST SELLERS WEEK "+txtWeekNo.Text.Trim();
                    xlSheetM.get_Range("A3", misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy") + " - " + DateTime.ParseExact(txtToDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy");
                    int j=0;
                    if (dtBestSeller.Rows.Count > 0)
                        j = WriteToExcelBestSeller(dtBestSeller, xlSheetM, 3);
                    else
                        j = 3;
                    
                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "BAH";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                        j = WriteToExcelBestSeller(dtBestSeller, xlSheetM, j);
                   

                   
                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetM,j);

                   
                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j =WriteToExcelBestSeller(dtBestSeller, xlSheetM,j);

                   
                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j =WriteToExcelBestSeller(dtBestSeller, xlSheetM,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j =WriteToExcelBestSeller(dtBestSeller, xlSheetM,j);

                    
                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j =WriteToExcelBestSeller(dtBestSeller, xlSheetM,j);

                    
                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j =WriteToExcelBestSeller(dtBestSeller, xlSheetM,j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j =WriteToExcelBestSeller(dtBestSeller, xlSheetM,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j =WriteToExcelBestSeller(dtBestSeller, xlSheetM,j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetM, j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetM, j);

                    
                    // For Ladieswear

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "BAH";
                    objBestSeller.DivisionCode = "L%";

                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    Excel1.Worksheet xlSheetL = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                    xlSheetL.Name = "Ladieswear";
                    xlSheetL.get_Range("A2", misValue).Formula = "LADIESWEAR BEST SELLERS WEEK " + txtWeekNo.Text.Trim();
                    xlSheetL.get_Range("A3", misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy") + " - " + DateTime.ParseExact(txtToDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy");
                    
                    if (dtBestSeller.Rows.Count > 0)
                        j = WriteToExcelBestSeller(dtBestSeller, xlSheetL, 3);
                    else
                        j = 3;
                    
                   
                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "BAH";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                        j = WriteToExcelBestSeller(dtBestSeller, xlSheetL, j);
                   

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetL,j);

                   
                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j= WriteToExcelBestSeller(dtBestSeller, xlSheetL,j);

                   
                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j= WriteToExcelBestSeller(dtBestSeller, xlSheetL,j);

                   
                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetL,j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetL,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetL,j);

                   
                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetL,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j= WriteToExcelBestSeller(dtBestSeller, xlSheetL,j);


                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetL, j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetL, j);



                    // For Childrenswear

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "BAH";
                    objBestSeller.DivisionCode = "C%";

                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    Excel1.Worksheet xlSheetC = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                    xlSheetC.Name = "Childrenswear";
                    xlSheetC.get_Range("A2", misValue).Formula = "CHILDRENSWEAR BEST SELLERS WEEK " + txtWeekNo.Text.Trim();
                    xlSheetC.get_Range("A3", misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy") + " - " + DateTime.ParseExact(txtToDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy");
                   
                    if (dtBestSeller.Rows.Count > 0)
                        j = WriteToExcelBestSeller(dtBestSeller, xlSheetC, 3);
                    else
                        j = 3;


                   
                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "BAH";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                        j = WriteToExcelBestSeller(dtBestSeller, xlSheetC, j);
                    
                    
                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetC,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetC,j);

                   
                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j =WriteToExcelBestSeller(dtBestSeller, xlSheetC,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetC,j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j =WriteToExcelBestSeller(dtBestSeller, xlSheetC,j);

                   
                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetC,j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetC,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetC,j);


                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetC, j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetC, j);

                    // For Footwear

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "BAH";
                    objBestSeller.DivisionCode = "F%";

                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    Excel1.Worksheet xlSheetF = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                    xlSheetF.Name = "Footwear And Accessories";
                    xlSheetF.get_Range("A2", misValue).Formula = "FOOTWEAR AND ACCESSORIES BEST SELLERS WEEK " + txtWeekNo.Text.Trim();
                    xlSheetF.get_Range("A3", misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy") + " - " + DateTime.ParseExact(txtToDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy");

                    if (dtBestSeller.Rows.Count > 0)
                        j = WriteToExcelBestSeller(dtBestSeller, xlSheetF, 3);
                    else
                        j = 3;
                   

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "BAH";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                      j = WriteToExcelBestSeller(dtBestSeller, xlSheetF, j);
                    
                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j =WriteToExcelBestSeller(dtBestSeller, xlSheetF,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetF,j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetF,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetF,j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetF,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetF,j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetF,j);

                   
                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetF,j);



                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetF, j);


                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetF, j);


                    // For Essentials

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "BAH";
                    objBestSeller.DivisionCode = "S%";

                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    Excel1.Worksheet xlSheetE = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                    xlSheetE.Name = "Own Brand Sports";
                    xlSheetE.get_Range("A2", misValue).Formula = "OWN BRAND SPORTS BEST SELLERS WEEK " + txtWeekNo.Text.Trim();
                    xlSheetE.get_Range("A3", misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy") + " - " + DateTime.ParseExact(txtToDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy");

                    if (dtBestSeller.Rows.Count > 0)
                        j = WriteToExcelBestSeller(dtBestSeller, xlSheetE, 3);
                    else
                        j = 3;
                    
                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "BAH";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                        j = WriteToExcelBestSeller(dtBestSeller, xlSheetE, j);
                    

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetE,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetE,j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetE,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetE,j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetE,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetE,j);


                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetE,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetE,j);


                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetE, j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    WriteToExcelBestSeller(dtBestSeller, xlSheetE, j);


                    // For Homeware

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "BAH";
                    objBestSeller.DivisionCode = "H%";

                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    Excel1.Worksheet xlSheetH = (Excel1.Worksheet)myExcelWorkbook.Sheets[6];
                    xlSheetH.Name = "Homeware";
                    xlSheetH.get_Range("A2", misValue).Formula = "HOMEWARE BEST SELLERS WEEK " + txtWeekNo.Text.Trim();
                    xlSheetH.get_Range("A3", misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy") + " - " + DateTime.ParseExact(txtToDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy");

                    if (dtBestSeller.Rows.Count > 0)
                        j = WriteToExcelBestSeller(dtBestSeller, xlSheetH, 3);
                    else
                        j = 3;
                    

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "BAH";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                     j = WriteToExcelBestSeller(dtBestSeller, xlSheetH, j);
                    

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetH, j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetH, j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetH, j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetH, j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetH, j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetH, j);


                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetH, j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetH, j);


                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetH, j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    WriteToExcelBestSeller(dtBestSeller, xlSheetH, j);


                    // For Summery

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "Summery";
                    objBestSeller.DivisionCode = "M%";

                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    Excel1.Worksheet xlSheetSummery = (Excel1.Worksheet)myExcelWorkbook.Sheets[7];
                    xlSheetSummery.Name = "Summary";
                    xlSheetSummery.get_Range("A2", misValue).Formula = "SUMMARY BEST SELLERS WEEK " + txtWeekNo.Text.Trim();
                    xlSheetSummery.get_Range("A3", misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy") + " - " + DateTime.ParseExact(txtToDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy");

                    if (dtBestSeller.Rows.Count > 0)
                        j = WriteToExcelBestSeller(dtBestSeller, xlSheetSummery, 3);
                    else
                        j = 3;

                   
                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.DivisionCode = "M%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetSummery, j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.DivisionCode="L%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetSummery, j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.DivisionCode = "L%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetSummery, j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.DivisionCode = "C%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetSummery, j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.DivisionCode = "C%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetSummery, j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.DivisionCode = "F%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetSummery, j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.DivisionCode = "F%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetSummery, j);


                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.DivisionCode = "S%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetSummery, j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.DivisionCode = "S%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetSummery, j);


                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.DivisionCode = "H%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetSummery, j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.DivisionCode = "H%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    WriteToExcelBestSeller(dtBestSeller, xlSheetSummery, j);

                }
                else
                {
                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "BAH";
                    objBestSeller.DivisionCode = ddlDivisionCode.SelectedItem.Value+"%";

                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    Excel1.Worksheet xlSheetOne = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                    xlSheetOne.Name = ddlDivisionCode.SelectedItem.Text;
                    xlSheetOne.get_Range("A2", misValue).Formula =  ddlDivisionCode.SelectedItem.Text.Trim().ToUpper()+" BEST SELLERS WEEK " + txtWeekNo.Text.Trim();
                    xlSheetOne.get_Range("A3", misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy") + " - " + DateTime.ParseExact(txtToDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy");
                    int j = 0;

                    if (dtBestSeller.Rows.Count > 0)
                        j = WriteToExcelBestSeller(dtBestSeller, xlSheetOne, 3);
                    else
                        j = 3;

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "BAH";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetOne,j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetOne,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j =WriteToExcelBestSeller(dtBestSeller, xlSheetOne,j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetOne,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetOne,j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetOne,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetOne,j);

                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetOne,j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j=WriteToExcelBestSeller(dtBestSeller, xlSheetOne,j);


                    objBestSeller.ReportType = "QtyWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetOne, j);

                    objBestSeller.ReportType = "SalesWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j = WriteToExcelBestSeller(dtBestSeller, xlSheetOne, j);
                }

            }
            else
            {
                //do nothing
            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\BestSellerReport_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

            ViewState["FileNameBestSeller"] = filePath;
            myExcelWorkbook.SaveAs(@filePath);

            myExcelWorkbook.Close();
            myExcelWorkbooks.Close();

            btnDownload.Visible = true;

            //}

            //catch (Exception e)
            //{

            //}

        }
        #endregion GenerateBestSellerReport

        #region Write To Excel Best Seller
        /// <summary>
        /// Write To Excel Best Seller
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private int WriteToExcelBestSeller(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, int j)
        {
            object misValue = System.Reflection.Missing.Value;
           // int j = 7;
            j = j + 4;

            if (dtStock.Rows.Count > 0)
            {
                string location = dtStock.Rows[0]["Location"].ToString();
                string ReportType = dtStock.Rows[0]["ReportType"].ToString();
                string Division = dtStock.Rows[0]["DivisionCode"].ToString();
                string DivisionCode="";
                  switch(Division)
                  {
                      case "M": DivisionCode = "Menswear";
                          break;
                      case "L": DivisionCode = "Ladieswear";
                          break;
                      case "C": DivisionCode = "Childrenswear";
                          break;
                      case "F": DivisionCode = "Footwear And Accessories";
                          break;
                      case "S": DivisionCode = "Own Brand Sports";
                          break;
                      case "H": DivisionCode = "Homeware";
                          break;

                  }

                switch (location)
                {
                    case "BAH":myExcelWorksheet.get_Range("A"+(j-2), misValue).Formula = "BAHRAIN  " +ReportType;
                               myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.Yellow;
                               break;
                    case "QAT": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Formula = "QATAR  " + ReportType;
                                myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.OrangeRed;
                               break;
                    case "OMAN": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Formula = "OMAN  " + ReportType;
                               myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.GreenYellow;
                               break;
                    case "JOR": myExcelWorksheet.get_Range("A" + (j -2), misValue).Formula = "JORDAN  " + ReportType;
                               myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.Gold;
                               break;
                    case "UAE": myExcelWorksheet.get_Range("A" + (j -2), misValue).Formula = "UAE  " + ReportType;
                               myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.MediumPurple;
                               break;
                    case "KSA": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Formula = "KSA  " + ReportType;
                               myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.AliceBlue;
                               break;

                    case "Summery": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Formula = DivisionCode+ " " + ReportType;
                               switch (Division)
                               {
                                   case "M": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.Yellow;
                                       break;
                                   case "L": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.Gold;
                                       break;
                                   case "C": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.MediumPurple;
                                       break;
                                   case "F": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.GreenYellow;
                                       break;
                                   case "S": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.OrangeRed;
                                       break;
                                   case "H": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.LightGray;
                                       break;
                               }
                               break;
                }

            }

            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {
                
                myExcelWorksheet.get_Range("A" + j, misValue).Formula = (null != dtStock.Rows[i]["ItemNo"]) ? dtStock.Rows[i]["ItemNo"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["Description"]) ? dtStock.Rows[i]["Description"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("C" + j, misValue).Formula = (null != dtStock.Rows[i]["ItemCategoryCode"]) ? dtStock.Rows[i]["ItemCategoryCode"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("C" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["SeasonCode"]) ? dtStock.Rows[i]["SeasonCode"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("D" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["Linecode7"]) ? dtStock.Rows[i]["Linecode7"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("E" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));


                myExcelWorksheet.get_Range("F" + j, misValue).Formula = (null != dtStock.Rows[i]["Qty"]) ? dtStock.Rows[i]["Qty"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("F" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("G" + j, misValue).Formula = (null != dtStock.Rows[i]["Sales"]) ? dtStock.Rows[i]["Sales"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("G" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
          
            }
            return j;
        }
        #endregion Write To Excel Best Seller



        #region GenerateBestSellerReportByLC7
        /// <summary>
        /// To generate excel report for Best Seller Report
        /// </summary>
        private void GenerateBestSellerReportByLC7()
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




            string fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\BestSellerReportLC7.xlsx";
            myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);



            //myExcelWorkbook = myExcelApp.Workbooks.Add(1);
            //Excel.Sheets xlSheets1 = myExcelWorkbook.Sheets as Excel.Sheets;

            //Excel.Worksheet xlSheet = (Excel.Worksheet)xlSheets1.Add(xlSheets1[1], Type.Missing, Type.Missing, Type.Missing);



            GetStockDetails objBestSeller = new GetStockDetails();

            //String cellFormulaAsString = myExcelWorksheet.get_Range("A2", misValue).Formula.ToString();
            if (txtFromDate.Text.Trim().Length > 0 && txtToDate.Text.Trim().Length > 0)
            {
                // string WeekNo = txtWeekNo.Text.Trim();

                fileName = HttpContext.Current.Server.MapPath(".") + "\\Template\\BestSellerReportLC7.xlsx";
                myExcelWorkbook = myExcelWorkbooks.Open(fileName, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                Excel1.Worksheet xlSheet = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];


                objBestSeller.FromDate = DateTime.ParseExact(txtFromDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                objBestSeller.ToDate = DateTime.ParseExact(txtToDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                objBestSeller.UaeRate = Convert.ToDecimal(txtUaeRate.Text.Trim());
                objBestSeller.BahRate = Convert.ToDecimal(txtBahrainRate.Text.Trim());
                objBestSeller.OmanRate = Convert.ToDecimal(txtOmanRate.Text.Trim());
                objBestSeller.JorRate = Convert.ToDecimal(txtJordanRate.Text.Trim());

                objBestSeller.KsaRate = Convert.ToDecimal(txtKsaRate.Text.Trim());

                string ReportType = ddlDivisionCode.SelectedItem.Value;
                if (ReportType == "All")
                {

                    // For Menswear

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "BAH";
                    objBestSeller.DivisionCode = "M%";

                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    Excel1.Worksheet xlSheetM = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                    xlSheetM.Name = "Menswear";
                    xlSheetM.get_Range("A2", misValue).Formula = "MENSWEAR BEST SELLERS WEEK " + txtWeekNo.Text.Trim();
                    xlSheetM.get_Range("A3", misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy") + " - " + DateTime.ParseExact(txtToDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy");
                    int j = 0;
                    if (dtBestSeller.Rows.Count > 0)
                        j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetM, 3);
                    else
                        j = 3;

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "BAH";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetM, j);



                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetM, j);


                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetM, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetM, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetM, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetM, j);


                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetM, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetM, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetM, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetM, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetM, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetM, j);


                    // For Ladieswear

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "BAH";
                    objBestSeller.DivisionCode = "L%";

                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    Excel1.Worksheet xlSheetL = (Excel1.Worksheet)myExcelWorkbook.Sheets[2];
                    xlSheetL.Name = "Ladieswear";
                    xlSheetL.get_Range("A2", misValue).Formula = "LADIESWEAR BEST SELLERS WEEK " + txtWeekNo.Text.Trim();
                    xlSheetL.get_Range("A3", misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy") + " - " + DateTime.ParseExact(txtToDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy");

                    if (dtBestSeller.Rows.Count > 0)
                        j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetL, 3);
                    else
                        j = 3;


                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "BAH";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetL, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetL, j);


                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetL, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetL, j);


                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetL, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetL, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetL, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetL, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetL, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetL, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetL, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetL, j);


                    // For Childrenswear

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "BAH";
                    objBestSeller.DivisionCode = "C%";

                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    Excel1.Worksheet xlSheetC = (Excel1.Worksheet)myExcelWorkbook.Sheets[3];
                    xlSheetC.Name = "Childrenswear";
                    xlSheetC.get_Range("A2", misValue).Formula = "CHILDRENSWEAR BEST SELLERS WEEK " + txtWeekNo.Text.Trim();
                    xlSheetC.get_Range("A3", misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy") + " - " + DateTime.ParseExact(txtToDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy");

                    if (dtBestSeller.Rows.Count > 0)
                        j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetC, 3);
                    else
                        j = 3;



                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "BAH";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetC, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetC, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetC, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetC, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetC, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetC, j);


                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetC, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetC, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetC, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetC, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetC, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetC, j);



                    // For Footwear

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "BAH";
                    objBestSeller.DivisionCode = "F%";

                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    Excel1.Worksheet xlSheetF = (Excel1.Worksheet)myExcelWorkbook.Sheets[4];
                    xlSheetF.Name = "Footwear And Accessories";
                    xlSheetF.get_Range("A2", misValue).Formula = "FOOTWEAR AND ACCESSORIES BEST SELLERS WEEK " + txtWeekNo.Text.Trim();
                    xlSheetF.get_Range("A3", misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy") + " - " + DateTime.ParseExact(txtToDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy");

                    if (dtBestSeller.Rows.Count > 0)
                        j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetF, 3);
                    else
                        j = 3;


                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "BAH";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetF, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetF, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetF, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetF, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetF, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetF, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetF, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetF, j);


                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetF, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetF, j);


                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetF, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetF, j);

                    // For Essentials

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "BAH";
                    objBestSeller.DivisionCode = "S%";

                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    Excel1.Worksheet xlSheetE = (Excel1.Worksheet)myExcelWorkbook.Sheets[5];
                    xlSheetE.Name = "Own Brand Sports";
                    xlSheetE.get_Range("A2", misValue).Formula = "OWN BRAND SPORTS BEST SELLERS WEEK " + txtWeekNo.Text.Trim();
                    xlSheetE.get_Range("A3", misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy") + " - " + DateTime.ParseExact(txtToDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy");

                    if (dtBestSeller.Rows.Count > 0)
                        j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetE, 3);
                    else
                        j = 3;

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "BAH";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetE, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetE, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetE, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetE, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetE, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetE, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetE, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetE, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j=WriteToExcelBestSellerLC7(dtBestSeller, xlSheetE, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetE, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetE, j);
                    WriteToExcelBestSellerLC7(dtBestSeller, xlSheetE, j);


                    // For Homeware

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "BAH";
                    objBestSeller.DivisionCode = "H%";

                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    Excel1.Worksheet xlSheetH = (Excel1.Worksheet)myExcelWorkbook.Sheets[6];
                    xlSheetH.Name = "Homeware";
                    xlSheetH.get_Range("A2", misValue).Formula = "HOMEWARE BEST SELLERS WEEK " + txtWeekNo.Text.Trim();
                    xlSheetH.get_Range("A3", misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy") + " - " + DateTime.ParseExact(txtToDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy");

                    if (dtBestSeller.Rows.Count > 0)
                        j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetH, 3);
                    else
                        j = 3;


                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "BAH";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetH, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetH, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetH, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetH, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetH, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetH, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetH, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetH, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                  j=  WriteToExcelBestSellerLC7(dtBestSeller, xlSheetH, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetH, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetH, j);
                    WriteToExcelBestSellerLC7(dtBestSeller, xlSheetH, j);


                    // For Summery

                    dtBestSeller = null;

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "SummeryLC7";
                    objBestSeller.DivisionCode = "M%";

                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    Excel1.Worksheet xlSheetSummery = (Excel1.Worksheet)myExcelWorkbook.Sheets[7];
                    xlSheetSummery.Name = "Summary";
                    xlSheetSummery.get_Range("A2", misValue).Formula = "SUMMARY BEST SELLERS WEEK " + txtWeekNo.Text.Trim();
                    xlSheetSummery.get_Range("A3", misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy") + " - " + DateTime.ParseExact(txtToDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy");

                    if (dtBestSeller.Rows.Count > 0)
                        j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetSummery, 3);
                    else
                        j = 3;

                    
                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.DivisionCode = "M%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetSummery, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.DivisionCode = "L%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetSummery, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.DivisionCode = "L%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetSummery, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.DivisionCode = "C%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetSummery, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.DivisionCode = "C%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetSummery, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.DivisionCode = "F%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetSummery, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.DivisionCode = "F%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetSummery, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.DivisionCode = "S%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetSummery, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.DivisionCode = "S%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetSummery, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.DivisionCode = "H%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetSummery, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.DivisionCode = "H%";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetSummery, j);
                    WriteToExcelBestSellerLC7(dtBestSeller, xlSheetSummery, j);

                }
                else
                {
                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "BAH";
                    objBestSeller.DivisionCode = ddlDivisionCode.SelectedItem.Value + "%";

                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    Excel1.Worksheet xlSheetOne = (Excel1.Worksheet)myExcelWorkbook.Sheets[1];
                    xlSheetOne.Name = ddlDivisionCode.SelectedItem.Text;
                    xlSheetOne.get_Range("A2", misValue).Formula = ddlDivisionCode.SelectedItem.Text.Trim().ToUpper() + " BEST SELLERS WEEK " + txtWeekNo.Text.Trim();
                    xlSheetOne.get_Range("A3", misValue).Formula = DateTime.ParseExact(txtFromDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy") + " - " + DateTime.ParseExact(txtToDate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM, yyyy");
                    int j = 0;

                    if (dtBestSeller.Rows.Count > 0)
                        j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetOne, 3);
                    else
                        j = 3;

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "BAH";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetOne, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetOne, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "JOR";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetOne, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetOne, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "UAE";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetOne, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetOne, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "OMAN";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetOne, j);

                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetOne, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "QAT";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetOne, j);


                    objBestSeller.ReportType = "LC7QtyWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetOne, j);

                    objBestSeller.ReportType = "LC7SalesWise";
                    objBestSeller.Location = "KSA";
                    dtBestSeller = objBestSeller.GetBestSellerReport();
                    if (dtBestSeller.Rows.Count > 0)
                        WriteHeader(xlSheetOne, j);
                    j = WriteToExcelBestSellerLC7(dtBestSeller, xlSheetOne, j);
                }

            }
            else
            {
                //do nothing
            }

            Random rnd = new Random();
            string filePath = Server.MapPath(".") + "\\Reports\\BestSellerReportByLinecode7_" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";

            ViewState["FileNameBestSellerLC7"] = filePath;
            myExcelWorkbook.SaveAs(@filePath);

            myExcelWorkbook.Close();
            myExcelWorkbooks.Close();

            btnBestSellerLC7.Visible = true;

            //}

            //catch (Exception e)
            //{

            //}

        }
        #endregion GenerateBestSellerReportByLC7

        #region WriteToExcelBestSellerLC7
        /// <summary>
        ///WriteToExcelBestSellerLC7
        /// </summary>
        /// <param name="dtStock"></param>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="location"></param>
        private int WriteToExcelBestSellerLC7(DataTable dtStock, Excel1.Worksheet myExcelWorksheet, int j)
        {
            object misValue = System.Reflection.Missing.Value;
            // int j = 7;
            j = j + 4;

            if (dtStock.Rows.Count > 0)
            {
                string location = dtStock.Rows[0]["Location"].ToString();
                string ReportType = dtStock.Rows[0]["ReportType"].ToString();
                string Division = dtStock.Rows[0]["DivisionCode"].ToString();
                string DivisionCode = "";
                switch (Division)
                {
                    case "M": DivisionCode = "Menswear";
                        break;
                    case "L": DivisionCode = "Ladieswear";
                        break;
                    case "C": DivisionCode = "Childrenswear";
                        break;
                    case "F": DivisionCode = "Footwear And Accessories";
                        break;
                    case "S": DivisionCode = "Own Brand Sports";
                        break;
                    case "H": DivisionCode = "Homeware";
                        break;

                }

                switch (location)
                {
                    case "BAH": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Formula = "BAHRAIN  " + ReportType.Substring(3, ReportType.Length - 3);
                        myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.Yellow;
                        break;
                    case "QAT": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Formula = "QATAR  " + ReportType.Substring(3, ReportType.Length - 3);
                        myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.OrangeRed;
                        break;
                    case "OMAN": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Formula = "OMAN  " + ReportType.Substring(3, ReportType.Length - 3);
                        myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.GreenYellow;
                        break;
                    case "JOR": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Formula = "JORDAN  " + ReportType.Substring(3, ReportType.Length - 3);
                        myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.Gold;
                        break;
                    case "UAE": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Formula = "UAE  " + ReportType.Substring(3, ReportType.Length - 3);
                        myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.MediumPurple;
                        break;

                    case "KSA": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Formula = "KSA  " + ReportType.Substring(3, ReportType.Length - 3);
                                myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.AliceBlue;
                                break;

                    case "Summery": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Formula = DivisionCode + " " + ReportType.Substring(3, ReportType.Length - 3);
                        switch (Division)
                        {
                            case "M": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.Yellow;
                                break;
                            case "L": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.Gold;
                                break;
                            case "C": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.MediumPurple;
                                break;
                            case "F": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.GreenYellow;
                                break;
                            case "S": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.OrangeRed;
                                break;
                            case "H": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.LightGray;
                                break;
                        }
                        break;

                    case "SummeryLC7": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Formula = DivisionCode + " " + ReportType.Substring(3, ReportType.Length-3);
                        switch (Division)
                        {
                            case "M": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.Yellow;
                                break;
                            case "L": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.Gold;
                                break;
                            case "C": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.MediumPurple;
                                break;
                            case "F": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.GreenYellow;
                                break;
                            case "S": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.OrangeRed;
                                break;
                            case "H": myExcelWorksheet.get_Range("A" + (j - 2), misValue).Interior.Color = System.Drawing.Color.LightGray;
                                break;
                        }
                        break;
                }

            }

            for (int i = 0; i < dtStock.Rows.Count; i++, j++)
            {

                myExcelWorksheet.get_Range("A" + j, misValue).Formula = (null != dtStock.Rows[i]["Linecode7"]) ? dtStock.Rows[i]["Linecode7"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("A" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("B" + j, misValue).Formula = (null != dtStock.Rows[i]["Description"]) ? dtStock.Rows[i]["Description"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("B" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("C" + j, misValue).Formula = (null != dtStock.Rows[i]["ItemCategoryCode"]) ? dtStock.Rows[i]["ItemCategoryCode"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("C" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("D" + j, misValue).Formula = (null != dtStock.Rows[i]["SeasonCode"]) ? dtStock.Rows[i]["SeasonCode"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("D" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));
                
                myExcelWorksheet.get_Range("E" + j, misValue).Formula = (null != dtStock.Rows[i]["Qty"]) ? dtStock.Rows[i]["Qty"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("E" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

                myExcelWorksheet.get_Range("F" + j, misValue).Formula = (null != dtStock.Rows[i]["Sales"]) ? dtStock.Rows[i]["Sales"].ToString() : "0";
                BorderAround(myExcelWorksheet.get_Range("F" + j, misValue), System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black));

            }
            return j;
        }
        #endregion WriteToExcelBestSellerLC7


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

        #region Write Header
        /// <summary>
        /// Write Header
        /// </summary>
        /// <param name="myExcelWorksheet"></param>
        /// <param name="j"></param>
        private void WriteHeader(Excel1.Worksheet myExcelWorksheet,int j)
        {
            Excel1.Range RngToCopy = myExcelWorksheet.get_Range("A5", "G6").EntireRow;
            Excel1.Range RngToInsert = myExcelWorksheet.get_Range("A" + (j +2), Type.Missing).EntireRow;
            RngToInsert.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, RngToCopy.Copy(Type.Missing));
        }
        #endregion Write Header

       
       
        #endregion Methods
    }
}