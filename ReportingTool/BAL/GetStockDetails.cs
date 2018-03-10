
#region NameSpace
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using Test.DAL;
#endregion NameSpace


namespace Test.BAL
{
    public class GetStockDetails
    {

        #region Public Properties
        /// <summary>
        /// Location code
        /// </summary>
        public string Location
        {
            get;
            set;
        }

        /// <summary>
        /// Posting Date
        /// </summary>
        public DateTime PostingDate
        {
            get;
            set;
        }

        /// <summary>
        /// From Date
        /// </summary>
        public DateTime FromDate
        {
            get;
            set;
        }

        /// <summary>
        /// To Date
        /// </summary>
        public DateTime ToDate
        {
            get;
            set;
        }

        /// <summary>
        /// ItemOperationType
        /// </summary>
        public Boolean ItemOperationType
        {
            get;
            set;
        }

        /// <summary>
        /// ILEOperationType
        /// </summary>
        public int ILEOperationType
        {
            get;
            set;
        }

        /// <summary>
        /// ValueOperationType
        /// </summary>
        public int ValueOperationType
        {
            get;
            set;
        }

        /// <summary>
        /// Foot Fall
        /// </summary>
        public Boolean FootFallOperationType
        {
            get;
            set;
        }

        /// <summary>
        /// Transaction Operation Type
        /// </summary>
        public Boolean TransactionOperationType
        {
            get;
            set;
        }

        /// <summary>
        /// SSOperationType
        /// </summary>
        public Boolean SSOperationType
        {
            get;
            set;
        }

        /// <summary>
        /// SSWeeklyOperationType
        /// </summary>
        public Boolean SSWeeklyOperationType
        {
            get;
            set;
        }

        /// <summary>
        /// SSReportOperationType
        /// </summary>
        public Boolean SSReportOperationType
        {
            get;
            set;
        }

        /// <summary>
        /// Uae Exchange Rate
        /// </summary>
        public decimal UaeRate
        {
            get;
            set;
        }

        /// <summary>
        /// Jordan Exchange Rate
        /// </summary>
        public decimal JorRate
        {
            get;
            set;
        }

        /// <summary>
        /// Oman Exchange Rate
        /// </summary>
        public decimal OmanRate
        {
            get;
            set;
        }

        /// <summary>
        /// Bahrain Exchange Rate
        /// </summary>
        public decimal BahRate
        {
            get;
            set;
        }

        /// <summary>
        /// KSA Exchange Rate
        /// </summary>
        public decimal KsaRate
        {
            get;
            set;
        }


        /// <summary>
        /// DtDataSource
        /// </summary>
        public DataTable DtSource
        {
            get;
            set;
        }

        /// <summary>
        /// Exception Message
        /// </summary>
        public string ExceptionMessage
        {
            get;
            set;
        }
        /// <summary>
        /// Week No
        /// </summary>
        public int WeekNo
        {
            get;
            set;
        }
        /// <summary>
        /// Year
        /// </summary>
        public string Year
        {
            get;
            set;
        }

        /// <summary>
        /// Integer Year
        /// </summary>
        public int IntYear
        {
            get;
            set;
        }
        /// <summary>
        /// Season Code
        /// </summary>
        public string SeasonCode
        {
            get;
            set;
        }

        /// <summary>
        /// Type
        /// </summary>
        public bool Type
        {
            get;
            set;
        }

        /// <summary>
        /// UaeOffer
        /// </summary>
        public string UaeOffer
        {
            get;
            set;
        }
        /// <summary>
        /// JordanOffer
        /// </summary>
        public string JordanOffer
        {
            get;
            set;
        }

        /// <summary>
        /// BahrainOffer
        /// </summary>
        public string BahrainOffer
        {
            get;
            set;
        }
        /// <summary>
        /// OmanOffer
        /// </summary>
        public string OmanOffer
        {
            get;
            set;
        }

        /// <summary>
        /// ProcessStatusId
        /// </summary>
        public int ProcessStatusId
        {
            get;
            set;
        }

        /// <summary>
        /// ProcessStatusFlag
        /// </summary>
        public bool ProcessStatusFlag
        {
            get;
            set;
        }

        /// <summary>
        /// Division Code
        /// </summary>
        public string DivisionCode
        {
            get;
            set;
        }

        /// <summary>
        /// Company Name
        /// </summary>
        public string CompanyName
        {
            get;
            set;
        }

        /// <summary>
        /// As Of Date
        /// </summary>
        public DateTime AsOfDate
        {
            get;
            set;
        }
        /// <summary>
        /// Report Type
        /// </summary>
        public string ReportType
        {
            get;
            set;
        }

        /// <summary>
        /// QatarOffer
        /// </summary>
        public string QatarOffer
        {
            get;
            set;
        }

        /// <summary>
        /// KsaOffer
        /// </summary>
        public string KsaOffer
        {
            get;
            set;
        }

        /// <summary>
        /// Country
        /// </summary>
        public string Country
        {
            get;
            set;
        }

        /// <summary>
        /// LYear
        /// </summary>
        public string LYear
        {
            get;
            set;
        }

        /// <summary>
        /// L2Year
        /// </summary>
        public string L2Year
        {
            get;
            set;
        }


        /// <summary>
        /// FromWeekNo
        /// </summary>
        public int FromWeekNo
        {
            get;
            set;
        }

        /// <summary>
        /// From Date LY
        /// </summary>
        public DateTime FromDateLY
        {
            get;
            set;
        }

        /// <summary>
        /// toDateLY
        /// </summary>
        public DateTime ToDateLY
        {
            get;
            set;
        }

        /// <summary>
        /// fromDate2LY
        /// </summary>
        public DateTime FromDate2LY
        {
            get;
            set;
        }

        /// <summary>
        /// toDate2LY
        /// </summary>
        public DateTime ToDate2LY
        {
            get;
            set;
        }

        /// <summary>
        ///fromDateYear
        /// </summary>
        public DateTime FromDateYear
        {
            get;
            set;
        }

        /// <summary>
        /// ToDate2LY
        /// </summary>
        public DateTime ToDateYear
        {
            get;
            set;
        }

        /// <summary>
        ///fromDateYearLY
        /// </summary>
        public DateTime FromDateYearLY
        {
            get;
            set;
        }

        /// <summary>
        /// toDate2LY
        /// </summary>
        public DateTime ToDateYearLY
        {
            get;
            set;
        }

        /// <summary>
        ///fromDateYear2LY
        /// </summary>
        public DateTime FromDateYear2LY
        {
            get;
            set;
        }

        /// <summary>
        /// toDate2LY
        /// </summary>
        public DateTime ToDateYear2LY
        {
            get;
            set;
        }

        /// <summary>
        /// ReportDate
        /// </summary>
        public DateTime ReportDate
        {
            get;
            set;
        }

        /// <summary>
        ///  Category Code
        /// </summary>
        public String CategoryCode
        {
            get;
            set;
        }

        /// <summary>
        ///  Line Code
        /// </summary>
        public String LineCode
        {
            get;
            set;
        }

        /// <summary>
        ///  PONumber
        /// </summary>
        public String PONumber
        {
            get;
            set;
        }

        /// <summary>
        ///  DocNumber
        /// </summary>
        public String DocNo
        {
            get;
            set;
        }




        /// <summary>
        ///  LineCode7
        /// </summary>
        public String LineCode7
        {
            get;
            set;
        }

        /// <summary>
        ///  PackID
        /// </summary>
        public String PackID
        {
            get;
            set;
        }

        /// <summary>
        ///  PackBarcode
        /// </summary>
        public String PackBarcode
        {
            get;
            set;
        }

        /// <summary>
        ///  PackType OrderQty
        /// </summary>
        public String PackType
        {
            get;
            set;
        }

        /// <summary>
        ///   OrderQty
        /// </summary>
        public decimal OrderQty
        {
            get;
            set;
        }


        /// <summary>
        ///   UnitPrice
        /// </summary>
        public decimal UnitPrice
        {
            get;
            set;
        }

        /// <summary>
        ///   COO
        /// </summary>
        public string COO
        {
            get;
            set;
        }

        /// <summary>
        ///   Department
        /// </summary>
        public string Department
        {
            get;
            set;
        }

        /// <summary>
        /// Nest
        /// </summary>
        public string Nest
        {
            get;
            set;
        }

        /// <summary>
        /// Description 
        /// </summary>
        public string Description
        {
            get;
            set;
        }

        /// <summary>
        /// Season 
        /// </summary>
        public string Season
        {
            get;
            set;
        }


        /// <summary>
        /// Outer 
        /// </summary>
        public decimal Outer
        {
            get;
            set;
        }

        /// <summary>
        /// Invoiced 
        /// </summary>
        public decimal Invoiced
        {
            get;
            set;
        }

        /// <summary>
        /// PackLevel 
        /// </summary>
        public string PackLevel
        {
            get;
            set;
        }

        /// <summary>
        /// ID 
        /// </summary>
        public int ID
        {
            get;
            set;
        }

        /// <summary>
        /// LineCode7Qty 
        /// </summary>
        public Decimal LineCode7Qty
        {
            get;
            set;
        }

        /// <summary>
        /// SalesAmount 
        /// </summary>
        public Decimal SalesAmount
        {
            get;
            set;
        }

        /// <summary>
        /// MonthNo 
        /// </summary>
        public int MonthNo
        {
            get;
            set;
        }

        /// <summary>
        ///  Month
        /// </summary>
        public String Month
        {
            get;
            set;
        }

        /// <summary>
        /// FromMonth
        /// </summary>
        public int FromMonth
        {
            get;
            set;
        }

        /// <summary>
        /// ToMonth
        /// </summary>
        public int ToMonth
        {
            get;
            set;
        }

        /// <summary>
        /// IntType
        /// </summary>
        public int IntType
        {
            get;
            set;
        }

        /// <summary>
        /// CevaIssueNo
        /// </summary>
        public string CevaIssueNo
        {
            get;
            set;
        }
        /// <summary>
        /// ItemNo
        /// </summary>
        public string ItemNo
        {
            get;
            set;
        }

        /// <summary>
        /// UnitCost
        /// </summary>
        public decimal UnitCost
        {
            get;
            set;
        }

        /// <summary>
        /// LineAmount
        /// </summary>
        public decimal LineAmount
        {
            get;
            set;
        }

        #endregion Public Properties


        #region GetAllStockValues
        /// <summary>
        /// Get All Stock Values
        /// </summary>
        /// <returns>Datatable Containing All Stock Values</returns>
        public DataTable GetAllStockValues(string locationCode)
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();
                //objStock.FromDate = fromDate;
                //objStock.ToDate = toDate;
                objStock.Location = locationCode;
                dtTest = objStock.GetStockValues();
               // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetAllStockValues

        #region GetAllStockValuesLCP
        /// <summary>
        /// Get All Stock Values LCP
        /// </summary>
        /// <returns>Datatable Containing All Stock Values LCP</returns>
        public DataTable GetAllStockValuesLCP(string locationCode)
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();
              
                objStock.Location = locationCode;
                dtTest = objStock.GetStockStatusLCP();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetAllStockValuesLCP

        #region GetWSSIReport
        /// <summary>
        /// GetWSSIReport
        /// </summary>
        /// <returns>Datatable Containing Get WSSI Report</returns>
        public DataTable GetWSSIReport()
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();
               
                objStock.WeekNo =WeekNo;
                objStock.Year = Year;
                objStock.SeasonCode = SeasonCode;
                dtTest = objStock.GetWSSIReport();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetWSSIReport

        #region GetWSSIForcast
        /// <summary>
        /// Get WSSI Forcast
        /// </summary>
        /// <returns>Datatable Containing  WSSI Forcast</returns>
        public DataTable GetWSSIForcast()
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.WeekNo = WeekNo;
                objStock.Year = Year;
                objStock.Type = Type;

                dtTest = objStock.GetWSSIForcast();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetWSSIForcast


        #region GetVisitorsReport
        /// <summary>
        /// Get Visitors Report
        /// </summary>
        /// <returns>Datatable Containing All Visitors Report</returns>
        public DataTable GetVisitorsReport(string locationCode)
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();
                //objStock.FromDate = fromDate;
                //objStock.ToDate = toDate;
                objStock.Location = locationCode;
                dtTest = objStock.GetVisitorsReport();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetVisitorsReport


        #region GetStockStatusLCPSummery
        /// <summary>
        /// Get Stock Status LCP Summery
        /// </summary>
        /// <returns>Datatable Containing Stock Status LCP Summery</returns>
        public DataTable GetStockStatusLCPSummery()
        {
            DataTable dtStock = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;

                dtStock = objStock.GetStockStatusLCPSummery();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtStock;
        }
        #endregion GetStockStatusLCPSummery

        #region GetPgCmpReport
        /// <summary>
        /// Get PgCmp Report
        /// </summary>
        /// <returns>Datatable Containing  PgCmpReport</returns>
        public DataTable GetPgCmpReport()
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Location =Location;
                dtTest = objStock.GetPgCmpReport();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetPgCmpReport



        #region InsertStockStatus
        /// <summary>
        /// Get All Stock Values
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertStockStatus()
        {
            bool Result=false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.SSOperationType = SSOperationType;
                objStock.SSWeeklyOperationType = SSWeeklyOperationType;
                objStock.SSReportOperationType = SSReportOperationType;

                objStock.JorRate = JorRate;
                objStock.UaeRate = UaeRate;
                objStock.BahRate = BahRate;
                objStock.OmanRate = OmanRate;

                objStock.KsaRate =KsaRate;
                Result = objStock.InsertStockStatus();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertStockStatus


        #region UpdateTables
        /// <summary>
        /// Update Tables
        /// </summary>
        /// <returns>Result</returns>
        public bool UpdateTables()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                
                objStock.ItemOperationType = ItemOperationType;
                objStock.ILEOperationType = ILEOperationType;
                objStock.ValueOperationType = ValueOperationType;

                objStock.FootFallOperationType = FootFallOperationType;
                objStock.TransactionOperationType = TransactionOperationType;

                Result = objStock.UpdateTables();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion UpdateTables


        #region UpdateOfferPrice
        /// <summary>
        /// Update Offer Price
        /// </summary>
        /// <returns>Result</returns>
        public bool UpdateOfferPrice()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.UaeOffer=UaeOffer;
                objStock.BahrainOffer=BahrainOffer;
                
                objStock.OmanOffer=OmanOffer;
                objStock.JordanOffer=JordanOffer;
                objStock.QatarOffer = QatarOffer;
                objStock.KsaOffer = KsaOffer;
                Result = objStock.UpdateOfferPrice();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion UpdateTables


        #region InsertVisitorsReport
        /// <summary>
        /// Insert Visitors Report
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertVisitorsReport()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.PostingDate = PostingDate;
               
                Result = objStock.InsertVisitorsReport();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertVisitorsReport

        #region BulkInsert
        /// <summary>
        /// Bulk Insert
        /// </summary>
        /// <returns>Result</returns>
        public bool BulkInsert()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.DtSource = DtSource;
                objStock.BultInsert();
                ExceptionMessage = objStock.ExceptionMessage;

                if (null!=ExceptionMessage && ExceptionMessage.ToString().Length > 0)
                {
                    Result = false;
                }
                else
                {
                    Result = true;
                }
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion BulkInsert

        #region GetVisitorsWeeklyReport
        /// <summary>
        /// Get Visitors Weekly Report
        /// </summary>
        /// <returns>Datatable Containing All Visitors Weekly Report</returns>
        public DataTable GetVisitorsWeeklyReport(string locationCode,DateTime PostingDate)
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objVisitors = new StockDetails();

                objVisitors.Location = locationCode;
                objVisitors.PostingDate = PostingDate;
                dtTest = objVisitors.GetVisitorsWeeklyReport();
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetVisitorsWeeklyReport

        #region InsertWssiReport
        /// <summary>
        /// Insert Wssi Report
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertWssiReport()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.WeekNo =WeekNo ;
                objStock.Year =Year;
                
                objStock.BahRate = BahRate;
                objStock.OmanRate = OmanRate;
                objStock.JorRate = JorRate;
                objStock.UaeRate = UaeRate;

                objStock.KsaRate = KsaRate;
                Result = objStock.InsertWssiReport();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertWssiReport

        #region InsertProductGroupCmpReport
        /// <summary>
        /// Insert Product Group Cmp Report
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertProductGroupCmpReport()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.WeekNo = WeekNo;
                objStock.IntYear = IntYear;

                objStock.BahRate = BahRate;
                objStock.JorRate = JorRate;
                objStock.OmanRate = OmanRate;
                objStock.UaeRate = UaeRate;

                objStock.KsaRate = KsaRate;

                Result = objStock.InsertProductGroupCmpReport();

                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertWssiReport


        #region GetProcessStatus
        /// <summary>
        /// Get Process Status
        /// </summary>
        /// <returns>Datatable Containing  ProcessStatus</returns>
        public DataTable GetProcessStatus()
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.ProcessStatusId = ProcessStatusId;
                dtTest = objStock.GetProcessStatus();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetProcessStatus


        #region UpdateProcessStatus
        /// <summary>
        /// Update Process Status
        /// </summary>
        /// <returns>Result</returns>
        public bool UpdateProcessStatus()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.ProcessStatusFlag = ProcessStatusFlag;
                objStock.ProcessStatusId = ProcessStatusId;

                Result = objStock.UpdateProcessStatus();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion UpdateTables

        #region Insert Wssi Division Report
        /// <summary>
        /// Insert Wssi Division Report
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertWssiDivisionReport()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.WeekNo = WeekNo;
                objStock.Year = Year;

                objStock.BahRate = BahRate;
                objStock.OmanRate = OmanRate;
                objStock.JorRate = JorRate;
                objStock.UaeRate = UaeRate;

                Result = objStock.InsertWssiDivisionReport();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Insert Wssi Division Report

        #region GetWSSIDivisionReport
        /// <summary>
        /// Get WSSI Division Report
        /// </summary>
        /// <returns>Datatable Containing Get WSSI Division Report</returns>
        public DataTable GetWSSIDivisionReport()
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.WeekNo = WeekNo;
                objStock.Year = Year;
                objStock.DivisionCode = DivisionCode;
                dtTest = objStock.GetWSSIDivisionReport();
            }
            catch (Exception e)
            {
                
            }
            return dtTest;
        }
        #endregion GetWSSIDivisionReport

        #region Insert Wssi Product Group Report
        /// <summary>
        /// Insert Wssi Product Group Report
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertWssiProductGroupReport()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;

                Result = objStock.InsertWssiProductGroupReport();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Insert Wssi Product Group Report

        #region GetWSSIProductGroupReport
        /// <summary>
        /// Get WSSI Product Group Report
        /// </summary>
        /// <returns>Datatable Containing Get WSSI Product Group Report</returns>
        public DataTable GetWSSIProductGroupReport()
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();

                dtTest = objStock.GetWSSIProductGroupReport();
            }
            catch (Exception e)
            {

            }
            return dtTest;
        }
        #endregion GetWSSIProductGroupReport


        //--------------------------------

        #region GetStockSummery
        /// <summary>
        /// Get Stock Summery
        /// </summary>
        /// <returns>Datatable Containing Get Stock Summery</returns>
        public DataTable GetStockSummary()
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.Location = Location;
                objStock.CompanyName =CompanyName;
                objStock.AsOfDate =AsOfDate;
                dtTest = objStock.GetStockSummary();
            }
            catch (Exception e)
            {

            }
            return dtTest;
        }
        #endregion GetStockSummery

        #region GetInventoryReport
        /// <summary>
        /// Get Inventory Report
        /// </summary>
        /// <returns>Datatable Containing Get Inventory Report</returns>
        public DataTable GetInventoryReport()
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                dtTest = objStock.GetInventoryReport();
            }
            catch (Exception e)
            {

            }
            return dtTest;
        }
        #endregion GetInventoryReport
        //--------------------------------

        // Best Seller Report -- Start

        #region Insert Best Seller Report
        /// <summary>
        /// Insert Best Seller Report
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertBestSellerReport()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.DivisionCode = DivisionCode;
                objStock.Location = Location;
                Result = objStock.InsertBestSellerReport();

                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Insert Best Seller Report

        #region Insert Best Seller By Linecode7
        /// <summary>
        /// Insert Best Seller By Linecode7
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertBestSellerByLinecode7()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.DivisionCode = DivisionCode;
                objStock.Location = Location;
                Result = objStock.InsertBestSellerByLinecode7();

                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Insert Best Seller By Linecode7

        #region Insert Best Seller Summery Report
        /// <summary>
        /// Insert Best Seller Summery Report
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertBestSellerSummeryReport()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.DivisionCode = DivisionCode;

                objStock.BahRate = BahRate;
                objStock.OmanRate = OmanRate;
                objStock.JorRate = JorRate;
                objStock.UaeRate = UaeRate;

                Result = objStock.InsertBestSellerSummeryReport();

                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Insert Best Seller Summery Report


        #region Insert Best Seller Summery Report LC7
        /// <summary>
        /// Insert Best Seller Summery Report LC7
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertBestSellerSummeryReportLC7()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.DivisionCode = DivisionCode;

                objStock.BahRate = BahRate;
                objStock.OmanRate = OmanRate;
                objStock.JorRate = JorRate;
                objStock.UaeRate = UaeRate;

                Result = objStock.InsertBestSellerSummeryReportLC7();

                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertBestSellerSummeryReportLC7

                

        #region Get Best Seller Report
        /// <summary>
        /// Get Best Seller Report
        /// </summary>
        /// <returns>Datatable Containing Best Seller Report</returns>
        public DataTable GetBestSellerReport()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.ReportType = ReportType;
                objStock.DivisionCode = DivisionCode;
                objStock.Location = Location;

                objStock.BahRate = BahRate;
                objStock.OmanRate = OmanRate;
                objStock.JorRate = JorRate;
                objStock.UaeRate = UaeRate;

                objStock.KsaRate = KsaRate;
                dtReport = objStock.GetBestSellerReport();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion Get Best Seller Report

        #region DeleteBestSellerReport
        /// <summary>
        /// DeleteBestSellerReport
        /// </summary>
        /// <returns>Result</returns>
        public void DeleteBestSellerReport()
        {
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.DeleteBestSellerReport();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
        }
        #endregion DeleteBestSellerReport
        // Best Seller Report -- End


        //DcStock Report- Start

        #region GetDCStock
        /// <summary>
        /// GetDCStock
        /// </summary>
        /// <returns>Datatable Containing DCStock</returns>
        public DataTable GetDCStock()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.AsOfDate = AsOfDate;
                objStock.Location = Location;
                objStock.Type = Type;
                
                dtReport = objStock.GetDCStock();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetDCStock
        //DcStock Report- End

        #region GetVisitorsVsSales
        /// <summary>
        /// GetVisitorsVsSales
        /// </summary>
        /// <returns>Datatable Containing VisitorsVsSales</returns>
        public DataTable GetVisitorsVsSales()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.PostingDate = PostingDate;
                objStock.Location = Location;
                dtReport = objStock.GetVisitorsVsSales();
            
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetVisitorsVsSales


        // Retail KPI Report

        #region GetRetailKPI
        /// <summary>
        /// GetRetailKPI
        /// </summary>
        /// <returns>Datatable Containing RetailKPI</returns>
        public DataTable GetRetailKPI()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Country =Country;
                dtReport = objStock.GetRetailKPI();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetRetailKPI


        #region Insert Retail KPI
        /// <summary>
        /// Insert Retail KPI
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertRetailKPI()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.WeekNo = WeekNo;
                objStock.Year = Year;
                objStock.LYear = LYear;
                objStock.L2Year = L2Year;
                
                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.BahRate = BahRate;
                objStock.OmanRate = OmanRate;
                
                objStock.JorRate = JorRate;
                objStock.UaeRate = UaeRate;
                objStock.KsaRate = KsaRate;
                objStock.ReportDate = ReportDate;

                Result = objStock.InsertRetailKPI();

                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Insert Retail KPI

        #region Insert Retail KPI By Division
        /// <summary>
        /// Insert Retail KPI By Division
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertRetailKPIByDivision()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.WeekNo = WeekNo;
                objStock.Year = Year;
                objStock.LYear = LYear;
                objStock.FromDate = FromDate;

                objStock.ToDate = ToDate;
                objStock.BahRate = BahRate;
                objStock.OmanRate = OmanRate;
                objStock.JorRate = JorRate;

                objStock.UaeRate = UaeRate;
                objStock.KsaRate = KsaRate;
                Result = objStock.InsertRetailKPIByDivision();

                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Insert Retail KPI By Division


        #region Get Retail KPI By Division
        /// <summary>
        /// Get Retail KPI By Division
        /// </summary>
        /// <returns>Datatable Containing Retail KPI By Division</returns>
        public DataTable GetRetailKPIByDivision()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Country = Country;
                dtReport = objStock.GetRetailKPIByDivision();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetRetailKPIByDivision


        #region Get Retail KPI LFL
        /// <summary>
        /// Get Retail KPI LFL
        /// </summary>
        /// <returns>Datatable Containing Get Retail KPI LFL</returns>
        public DataTable GetRetailKPILFL()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Country = Country;
                dtReport = objStock.GetRetailKPILFL();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion Get Retail KPI LFL


        #region GetRetailKPIByDivisionLFL
        /// <summary>
        ///GetRetailKPIByDivisionLFL
        /// </summary>
        /// <returns>Datatable Containing Get Retail KPI By Division LFL</returns>
        public DataTable GetRetailKPIByDivisionLFL()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Country = Country;
                dtReport = objStock.GetRetailKPIByDivisionLFL();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetRetailKPIByDivisionLFL

        #region InsertWeeklySales
        /// <summary>
        /// Insert Weekly Sales
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertWeeklySales()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.WeekNo = WeekNo;
                objStock.Year = Year;
                Result = objStock.InsertWeeklySales();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Insert Weekly Sales


        #region ImportSalesPlan
        /// <summary>
        /// Import Sales Plan
        /// </summary>
        /// <returns>Result</returns>
        public bool ImportSalesPlan()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.DtSource = DtSource;
                Result=objStock.ImportSalesPlan();
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Import Sales Plan

        #region ImportLinearCount
        /// <summary>
        /// Import Linear Count
        /// </summary>
        /// <returns>Result</returns>
        public bool ImportLinearCount()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.DtSource = DtSource;
                Result = objStock.ImportLinearCount();
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Import Linear Count

        #region InsertRetailKpiMonth
        /// <summary>
        /// InsertRetailKpiMonth
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertRetailKpiMonth()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                
                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.FromDateLY = FromDateLY;
                objStock.ToDateLY = ToDateLY;

                objStock.FromDate2LY = FromDate2LY;
                objStock.ToDate2LY = ToDate2LY;
                objStock.FromDateYear = FromDateYear;
                objStock.ToDateYear = ToDateYear;

                objStock.FromDateYearLY = FromDateYearLY;
                objStock.ToDateYearLY = ToDateYearLY;
                objStock.FromDateYear2LY = FromDateYear2LY;
                objStock.ToDateYear2LY = ToDateYear2LY;

                objStock.OmanRate = OmanRate;
                objStock.UaeRate = UaeRate;
                objStock.JorRate = JorRate;
                objStock.BahRate = BahRate;
                
                Result = objStock.InsertRetailKpiMonth();

                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertRetailKpiMonth

        #region InsertRetailKPIYearByDivision
        /// <summary>
        /// Insert RetailKPI Year By Division
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertRetailKPIYearByDivision()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.FromDateYear = FromDateYear;
                objStock.ToDateYear = ToDateYear;
                objStock.FromDateYearLY = FromDateYearLY;
                objStock.ToDateYearLY = ToDateYearLY;

                
                objStock.OmanRate = OmanRate;
                objStock.UaeRate = UaeRate;
                objStock.JorRate = JorRate;
                objStock.BahRate = BahRate;

                Result = objStock.InsertRetailKPIYearByDivision();

                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Insert RetailKPIYearByDivision

        #region InsertDailySales
        /// <summary>
        /// Insert Daily Sales
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertDailySales()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.Country = Country;
                Result = objStock.InsertDailySales();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Insert Daily Sales



        #region GetRetailKpiMonth
        /// <summary>
        /// GetRetailKpiMonth
        /// </summary>
        /// <returns>Datatable Containing RetailKpiMonth</returns>
        public DataTable GetRetailKpiMonth()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Country = Country;
                dtReport = objStock.GetRetailKpiMonth();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetRetailKpiMonth

        #region GetRetailKpiYear
        /// <summary>
        /// GetRetailKpiYear
        /// </summary>
        /// <returns>Datatable Containing RetailKpiYear</returns>
        public DataTable GetRetailKpiYear()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Country = Country;
                dtReport = objStock.GetRetailKpiYear(); 
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetRetailKpiYear

        #region GetRetailKpiYearDivision
        /// <summary>
        /// GetRetailKpiYearDivision
        /// </summary>
        /// <returns>Datatable Containing RetailKpiYearDivision</returns>
        public DataTable GetRetailKpiYearDivision()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Country = Country;
                dtReport = objStock.GetRetailKPIYearByDivision();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetRetailKpiYearDivision


        #region Delete Sales Plan
        /// <summary>
        /// Delete Sales Plan
        /// </summary>
        /// <returns>Result</returns>
        public bool DeleteSalesPlan()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.PostingDate = PostingDate;
                objStock.Location = Location;
                objStock.WeekNo = WeekNo;
                
                Result = objStock.DeleteSalesPlan();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Delete Sales Plan



        #region Delete Linear Count
        /// <summary>
        /// Delete Linear Count
        /// </summary>
        /// <returns>Result</returns>
        public bool DeleteLinearCount()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Location = Location;
                objStock.WeekNo = WeekNo;
                objStock.Year = Year;
                objStock.CategoryCode = CategoryCode;

                Result = objStock.DeleteLinearCount();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Delete Linear Count


        #region GetWeekDetails
        /// <summary>
        /// GetWeekDetails
        /// </summary>
        /// <returns>Datatable Containing WeekDetails</returns>
        public DataTable GetWeekDetails()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                dtReport = objStock.GetWeekDetails();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetWeekDetails

        // Retail KPI Report



        //Dsr Report Start
        #region InsertDsrReport
        /// <summary>
        /// Insert Dsr Report
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertDsrReport()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.FromDate = FromDate;
                objStock.Location= Location;
                
                objStock.JorRate = JorRate;
                objStock.KsaRate = KsaRate;
                objStock.UaeRate = UaeRate;
                objStock.OmanRate = OmanRate;

                objStock.BahRate = BahRate;
                Result = objStock.InsertDsrReport();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Insert Dsr Report


        #region GetDsrReport
        /// <summary>
        /// GetDsrReport
        /// </summary>
        /// <returns>Datatable Containing DsrReport</returns>
        public DataTable GetDsrReport()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Location = Location;
                dtReport = objStock.GetDsrReport();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetDsrReport


        #region InsertDsrDivision
        /// <summary>
        /// InsertDsrDivision
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertDsrDivision()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.FromDate = FromDate;
                objStock.ToDate= ToDate;
                objStock.JorRate = JorRate;
                objStock.KsaRate = KsaRate;
                
                objStock.UaeRate = UaeRate;
                objStock.OmanRate = OmanRate;
                objStock.BahRate = BahRate;

                Result = objStock.InsertDsrDivision();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Insert Dsr Report


        #region GetDsrDivision
        /// <summary>
        /// GetDsrDivision
        /// </summary>
        /// <returns>Datatable Containing GetDsrDivision</returns>
        public DataTable GetDsrDivision()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Location = Location;
                dtReport = objStock.GetDsrDivision();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetDsrDivision

        //Dsr Report End

        //Customer Count Report Start

        #region GetCustomerCount
        /// <summary>
        /// GetCustomerCount
        /// </summary>
        /// <returns>Datatable Containing CustomerCount</returns>
        public DataTable GetCustomerCount()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.Location = Location;
                dtReport = objStock.GetCustomerCount();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetCustomerCount

        
        //Customer Count Report End


        #region GetHighestClosingValues
        /// <summary>
        /// GetCustomerCount
        /// </summary>
        /// <returns>Datatable Containing CustomerCount</returns>
        public DataTable GetHighestClosingValues()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.PostingDate = PostingDate;
                objStock.ReportType =ReportType;
                objStock.Location = Location;
                dtReport = objStock.GetHighestClosingValues();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetHighestClosingValues


        #region GetPgcmpSummaryByDivision
        /// <summary>
        /// GetPgcmpSummaryByDivision
        /// </summary>
        /// <returns>Datatable Containing PgcmpSummaryByDivision</returns>
        public DataTable GetPgcmpSummaryByDivision()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                dtReport = objStock.GetPgcmpSummaryByDivision();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetPgcmpSummaryByDivision

        #region Get Item Info
        /// <summary>
        /// Get Item Info
        /// </summary>
        /// <returns>Datatable Containing GetItemInfo</returns>
        public DataTable GetItemInfo()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Country = Country;
                objStock.LineCode = LineCode;
                objStock.Location = Location;
                dtReport = objStock.GetItemInfo();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion Get Item Info

        #region Get Store By Country
        /// <summary>
        /// Get Store By Country
        /// </summary>
        /// <returns>Datatable Containing Get Store By Country</returns>
        public DataTable GetStoreByCountry()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Country = Country;
                dtReport = objStock.GetStoreByCountry();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion Get Store By Country


        #region InsertValueEntry
        /// <summary>
        /// InsertValueEntry
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertValueEntry()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.PostingDate = PostingDate;
                objStock.SalesAmount = SalesAmount;
                objStock.Description = Description;
                objStock.Location = Location;

                Result = objStock.InsertValueEntry();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
       #endregion InsertValueEntry





        //TATI Starts

        #region InsertWeeklySalesTati
        /// <summary>
        /// Insert Weekly Sales Tati
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertWeeklySalesTati()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.WeekNo = WeekNo;
                objStock.Year = Year;
                Result = objStock.InsertWeeklySalesTati();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertWeeklySalesTati

        #region InsertDailySalesTati
        /// <summary>
        /// Insert Daily Sales Tati
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertDailySalesTati()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.Country = Country;
                Result = objStock.InsertDailySalesTati();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertDailySalesTati


        #region Insert Retail KPI Tati
        /// <summary>
        /// Insert Retail KPI Tati
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertRetailKPITati()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.WeekNo = WeekNo;
                objStock.Year = Year;
                objStock.LYear = LYear;
                objStock.L2Year = L2Year;

                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.BahRate = BahRate;
                objStock.OmanRate = OmanRate;

                objStock.JorRate = JorRate;
                objStock.UaeRate = UaeRate;
                objStock.KsaRate = KsaRate;
                objStock.ReportDate = ReportDate;

                Result = objStock.InsertRetailKPITati();

                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Insert Retail KPI Tati

        
        #region Insert Retail KPI By Division Tati
        /// <summary>
        /// Insert Retail KPI By Division Tati
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertRetailKPIByDivisionTati()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.WeekNo = WeekNo;
                objStock.Year = Year;
                objStock.LYear = LYear;
                objStock.FromDate = FromDate;

                objStock.ToDate = ToDate;
                objStock.BahRate = BahRate;
                objStock.OmanRate = OmanRate;
                objStock.JorRate = JorRate;

                objStock.UaeRate = UaeRate;
                objStock.KsaRate = KsaRate;
                Result = objStock.InsertRetailKPIByDivisionTati();

                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Insert Retail KPI By Division Tati

        #region InsertDsrReportTati
        /// <summary>
        /// InsertDsrReportTati
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertDsrReportTati()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.FromDate = FromDate;
                objStock.Location = Location;

                objStock.JorRate = JorRate;
                objStock.KsaRate = KsaRate;
                objStock.UaeRate = UaeRate;
                objStock.OmanRate = OmanRate;

                objStock.BahRate = BahRate;
                Result = objStock.InsertDsrReportTati();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertDsrReportTati

        #region InsertDsrDivisionTati
        /// <summary>
        /// InsertDsrDivisionTati
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertDsrDivisionTati()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.JorRate = JorRate;
                objStock.KsaRate = KsaRate;

                objStock.UaeRate = UaeRate;
                objStock.OmanRate = OmanRate;
                objStock.BahRate = BahRate;

                Result = objStock.InsertDsrDivisionTati();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertDsrDivisionTati


        #region GetRetailKPITati
        /// <summary>
        /// GetRetailKPITati
        /// </summary>
        /// <returns>Datatable Containing RetailKPI</returns>
        public DataTable GetRetailKPITati()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Country = Country;
                dtReport = objStock.GetRetailKPITati();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetRetailKPITati

        #region Get Retail KPI LFL Tati
        /// <summary>
        /// Get Retail KPI LFL Tati
        /// </summary>
        /// <returns>Datatable Containing Get Retail KPI LFL Tati</returns>
        public DataTable GetRetailKPILFLTati()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Country = Country;
                dtReport = objStock.GetRetailKPILFLTati();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion Get Retail KPI LFL Tati

        
        #region Get Retail KPI By Division Tati
        /// <summary>
        /// Get Retail KPI By Division Tati
        /// </summary>
        /// <returns>Datatable Containing Retail KPI By Division</returns>
        public DataTable GetRetailKPIByDivisionTati()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Country = Country;
                dtReport = objStock.GetRetailKPIByDivisionTati();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion Get Retail KPI By Division Tati

        #region Get RetailKPI By Division LFL Tati
        /// <summary>
        ///GetRetailKPIByDivisionLFLTati
        /// </summary>
        /// <returns>Datatable Containing GetRetailKPIByDivisionLFLTati</returns>
        public DataTable GetRetailKPIByDivisionLFLTati()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Country = Country;
                dtReport = objStock.GetRetailKPIByDivisionLFLTati();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion Get RetailKPI By Division LFL Tati

        #region GetWeekDetailsTati
        /// <summary>
        /// GetWeekDetailsTati
        /// </summary>
        /// <returns>Datatable Containing WeekDetails</returns>
        public DataTable GetWeekDetailsTati()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                dtReport = objStock.GetWeekDetailsTati();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetWeekDetailsTati

        #region GetDsrReportTati
        /// <summary>
        /// GetDsrReportTati
        /// </summary>
        /// <returns>Datatable Containing GetDsrReportTati</returns>
        public DataTable GetDsrReportTati()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Location = Location;
                dtReport = objStock.GetDsrReportTati();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetDsrReportTati


        #region GetDsrDivisionTati
        /// <summary>
        /// GetDsrDivisionTati
        /// </summary>
        /// <returns>Datatable Containing GetDsrDivisionTati</returns>
        public DataTable GetDsrDivisionTati()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Location = Location;
                dtReport = objStock.GetDsrDivisionTati();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetDsrDivisionTati


        #region Delete Sales Plan Tati
        /// <summary>
        /// Delete Sales Plan Tati
        /// </summary>
        /// <returns>Result</returns>
        public bool DeleteSalesPlanTati()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.PostingDate = PostingDate;
                objStock.Location = Location;
                objStock.WeekNo = WeekNo;

                Result = objStock.DeleteSalesPlanTati();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Delete Sales Plan Tati


        #region ImportSalesPlanTati
        /// <summary>
        /// ImportSalesPlanTati
        /// </summary>
        /// <returns>Result</returns>
        public bool ImportSalesPlanTati()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.DtSource = DtSource;
                Result = objStock.ImportSalesPlanTati();
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion ImportSalesPlanTati

        #region Delete Linear Count Tati
        /// <summary>
        /// Delete Linear Count Tati
        /// </summary>
        /// <returns>Result</returns>
        public bool DeleteLinearCountTati()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Location = Location;
                objStock.WeekNo = WeekNo;
                objStock.Year = Year;
                objStock.CategoryCode = CategoryCode;

                Result = objStock.DeleteLinearCountTati();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion Delete Linear Count Tati

        #region ImportLinearCountTati
        /// <summary>
        /// ImportLinearCountTati
        /// </summary>
        /// <returns>Result</returns>
        public bool ImportLinearCountTati()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.DtSource = DtSource;
                Result = objStock.ImportLinearCountTati();
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion ImportLinearCountTati


        #region InsertVisitorsReportTati
        /// <summary>
        ///Insert Visitors Report Tati
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertVisitorsReportTati()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.PostingDate = PostingDate;

                Result = objStock.InsertVisitorsReportTati();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertVisitorsReportTati

        #region GetVisitorsReportTati
        /// <summary>
        /// GetVisitorsReportTati
        /// </summary>
        /// <returns>Datatable Containing All VisitorsReportTati</returns>
        public DataTable GetVisitorsReportTati(string locationCode)
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();
                //objStock.FromDate = fromDate;
                //objStock.ToDate = toDate;
                objStock.Location = locationCode;
                dtTest = objStock.GetVisitorsReportTati();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetVisitorsReportTati


        #region GetVisitorsWeeklyReportTati
        /// <summary>
        /// Get Visitors Weekly Report Tati
        /// </summary>
        /// <returns>Datatable Containing All Visitors Weekly Report</returns>
        public DataTable GetVisitorsWeeklyReportTati(string locationCode, DateTime PostingDate)
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objVisitors = new StockDetails();

                objVisitors.Location = locationCode;
                objVisitors.PostingDate = PostingDate;
                dtTest = objVisitors.GetVisitorsWeeklyReportTati();
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetVisitorsWeeklyReportTati

        #region InsertVisitorDataTati
        /// <summary>
        /// Insert Visitor Data Tati
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertVisitorDataTati()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.DtSource = DtSource;
                objStock.InsertVisitorDataTati();
                ExceptionMessage = objStock.ExceptionMessage;

                if (null != ExceptionMessage && ExceptionMessage.ToString().Length > 0)
                {
                    Result = false;
                }
                else
                {
                    Result = true;
                }
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertVisitorDataTati

        #region UpdateTablesTati
        /// <summary>
        /// Update Tables Tati
        /// </summary>
        /// <returns>Result</returns>
        public bool UpdateTablesTati()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.ItemOperationType = ItemOperationType;
                objStock.ILEOperationType = ILEOperationType;
                objStock.ValueOperationType = ValueOperationType;

                objStock.FootFallOperationType = FootFallOperationType;
                objStock.TransactionOperationType = TransactionOperationType;

                objStock.UpdateTablesMY();
                Result = objStock.UpdateTablesTati();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion UpdateTablesTati
        //TATI Ends


        //DCStock Start



        #region GetPOHeader
        /// <summary>
        ///GetPOHeader
        /// </summary>
        /// <returns>Datatable Containing GetPOHeader</returns>
        public DataTable GetPOHeader()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                dtReport = objStock.GetPOHeader();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetPOHeader


        #region GetPODetail
        /// <summary>
        ///GetPODetail
        /// </summary>
        /// <returns>Datatable Containing GetPODetail</returns>
        public DataTable GetPODetail()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.PONumber = PONumber;
                dtReport = objStock.GetPODetail();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetPODetail

        #region GetSODetail
        /// <summary>
        ///GetSODetail
        /// </summary>
        /// <returns>Datatable Containing GetSODetail</returns>
        public DataTable GetSODetail()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.DocNo = DocNo;
                dtReport = objStock.GetSODetail();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetSODetail


        #region GetSOHeader
        /// <summary>
        ///GetSOHeader
        /// </summary>
        /// <returns>Datatable Containing GetPOHeader</returns>
        public DataTable GetSOHeader()
        {
            DataTable dtReport = null;
            try
            {
                StockDetails objStock = new StockDetails();
                dtReport = objStock.GetSOHeader();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion GetSOHeader

        #region UpdatePODetail
        /// <summary>
        ///UpdatePODetail
        /// </summary>
        /// <returns>Result</returns>
        public bool UpdatePODetail()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.LineCode7 = LineCode7;
                objStock.PackID = PackID;
                objStock.PackBarcode = PackBarcode;
                objStock.PackType = PackType;

                objStock.OrderQty = OrderQty;
                objStock.UnitPrice = UnitPrice;
                objStock.COO = COO;
                objStock.Department = Department;

                objStock.Nest = Nest;
                objStock.Description = Description;
                objStock.Season = Season;
                objStock.Outer = Outer;

                objStock.Invoiced = Invoiced;
                objStock.PackLevel = PackLevel;
                objStock.ID = ID;
                Result = objStock.UpdatePODetail();
            }
            catch (Exception e)
            {

            }
            return Result;
        }
        #endregion UpdatePODetail


        #region InsertPODetail
        /// <summary>
        ///InsertPODetail
        /// </summary>
        /// <returns> Result</returns>
        public bool InsertPODetail()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.PONumber = PONumber;
                objStock.LineCode7 = LineCode7;
                objStock.PackID = PackID;
                objStock.PackBarcode = PackBarcode;

                objStock.PackType = PackType;
                objStock.OrderQty = OrderQty;
                objStock.UnitPrice = UnitPrice;
                objStock.COO = COO;

                objStock.Department = Department;
                objStock.Nest = Nest;
                objStock.Description = Description;
                objStock.Season = Season;

                objStock.Outer = Outer;
                objStock.Invoiced = Invoiced;
                objStock.PackLevel = PackLevel;
             
                Result = objStock.InsertPODetail();
            }
            catch (Exception e)
            {

            }
            return Result;
        }
        #endregion InsertPODetail


        #region DeletePODetail
        /// <summary>
        ///DeletePODetail
        /// </summary>
        /// <returns> Result</returns>
        public bool DeletePODetail()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.PONumber = PONumber;
                Result = objStock.DeletePODetail();
            }
            catch (Exception e)
            {

            }
            return Result;
        }
        #endregion DeletePODetail

        #region DeleteSODetail
        /// <summary>
        ///Delete SODetail
        /// </summary>
        /// <returns> Result</returns>
        public bool DeleteSODetail()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.DocNo = DocNo;
                Result = objStock.DeleteSODetail();
            }
            catch (Exception e)
            {

            }
            return Result;
        }
        #endregion DeleteSODetail



        #region UpdateStockLedger
        /// <summary>
        ///UpdateStockLedger
        /// </summary>
        /// <returns> Result</returns>
        public bool UpdateStockLedger()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.PackBarcode = PackBarcode;
                objStock.LineCode7 = LineCode7;
                objStock.PackID = PackID;
                objStock.PackType = PackType;

                objStock.LineCode7Qty = LineCode7Qty;
                objStock.Outer = Outer;
                objStock.PackLevel = PackLevel;
                objStock.ID = ID;
                
                Result = objStock.UpdateStockLedger();
            }
            catch (Exception e)
            {

            }
            return Result;
        }
        #endregion UpdateStockLedger


        //DCStock End


        //P&L Start


        #region GetProfitAndLoss
        /// <summary>
        /// GetProfitAndLoss
        /// </summary>
        /// <returns>Datatable Containing All GetProfitAndLoss</returns>
        public DataTable GetProfitAndLoss()
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.OmanRate = OmanRate;
                objStock.UaeRate = UaeRate;
                objStock.KsaRate = KsaRate;

                objStock.JorRate = JorRate;
                objStock.BahRate = BahRate;
                objStock.Location = Location;
                dtTest = objStock.GetProfitLoss();

                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetProfitAndLoss


        #region InsertProfitAndLossReport
        /// <summary>
        /// InsertProfitAndLossReport
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertProfitAndLossReport()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.MonthNo = MonthNo;
                objStock.Year = Year;
                Result = objStock.InsertProfitLossReport();
                ExceptionMessage = objStock.ExceptionMessage;

                if (null != ExceptionMessage && ExceptionMessage.ToString().Length > 0)
                {
                    Result = false;
                }
                else
                {
                    Result = true;
                }
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertProfitAndLossReport

        //P&L End


        //P&L Dynamic Start

        #region GetStoreDetails
        /// <summary>
        /// GetStoreDetails
        /// </summary>
        /// <returns>Datatable Containing All StoreDetails</returns>
        public DataTable GetStoreDetails()
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.IntType = IntType;
                dtTest = objStock.GetStoreDetails();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetStoreDetails

        #region InsertGLAccountDetails
        /// <summary>
        /// InsertGLAccountDetails
        /// </summary>
        /// <returns>Datatable Containing InsertGLAccountDetails</returns>
        public bool InsertGLAccountDetails()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Location = Location;
                objStock.Country = Country;
                Result = objStock.InsertGLAccountDetails();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertGLAccountDetails



        #region GetMonthDetails
        /// <summary>
        /// GetMonthDetails
        /// </summary>
        /// <returns>Datatable Containing All MonthDetails</returns>
        public DataTable GetMonthDetails()
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.Year = Year;
                objStock.FromMonth = FromMonth;
                objStock.ToMonth = ToMonth;
                dtTest = objStock.GetMonthDetails();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetMonthDetails

        #region UpdateProfitLossActualReport
        /// <summary>
        /// UpdateProfitLossActualReport
        /// </summary>
        /// <returns>Datatable Containing InsertGLAccountDetails</returns>
        public bool UpdateProfitLossActualReport()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.Month = Month;
                objStock.Location = Location;
                objStock.Country = Country;
                Result = objStock.UpdateProfitLossActualReport();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion UpdateProfitLossActualReport


        #region ImportProfitLoseBudgets
        /// <summary>
        /// ImportProfitLoseBudgets
        /// </summary>
        /// <returns>Datatable Containing ImportProfitLoseBudgets</returns>
        public bool ImportProfitLoseBudgets()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                Result = objStock.ImportProfitLoseBudgets();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion ImportProfitLoseBudgets


        #region UpdateProfitLossBudgetReport
        /// <summary>
        /// UpdateProfitLossBudgetReport
        /// </summary>
        /// <returns> UpdateProfitLossBudgetReport</returns>
        public bool UpdateProfitLossBudgetReport()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.Month = Month;
                objStock.Location = Location;

                objStock.Year = Year;
                Result = objStock.UpdateProfitLossBudgetReport();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion UpdateProfitLossBudgetReport


        #region DeleteProfitLossReport
        /// <summary>
        /// DeleteProfitLossReport
        /// </summary>
        /// <returns> DeleteProfitLossReport</returns>
        public bool DeleteProfitLossReport()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                Result = objStock.DeleteProfitLossReport();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion DeleteProfitLossReport



        //P&L Dynamic End


        //Stock Status VAT Start

        #region InsertStockStatusVAT
        /// <summary>
        ///InsertStockStatusVAT
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertStockStatusVAT()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.SSOperationType = SSOperationType;
                objStock.SSWeeklyOperationType = SSWeeklyOperationType;
                objStock.SSReportOperationType = SSReportOperationType;

                objStock.JorRate = JorRate;
                objStock.UaeRate = UaeRate;
                objStock.BahRate = BahRate;
                objStock.OmanRate = OmanRate;

                objStock.KsaRate = KsaRate;
                Result = objStock.InsertStockStatusVAT();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertStockStatusVAT

        #region GetStockStatusReportVAT
        /// <summary>
        /// Get All Stock Values
        /// </summary>
        /// <returns>Datatable Containing All Stock Values</returns>
        public DataTable GetStockStatusReportVAT(string locationCode)
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();
                //objStock.FromDate = fromDate;
                //objStock.ToDate = toDate;
                objStock.Location = locationCode;
                dtTest = objStock.GetStockStatusReportVAT();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetStockStatusReportVAT


        #region GetAllStockValuesLCPVAT
        /// <summary>
        /// Get All Stock Values LCP VAT
        /// </summary>
        /// <returns>Datatable Containing GetAllStockValuesLCPVAT</returns>
        public DataTable GetAllStockValuesLCPVAT(string locationCode)
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.Location = locationCode;
                dtTest = objStock.GetStockStatusLCPVAT();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetAllStockValuesLCPVAT


        //Stock Status VAT End



        //Invoice Report Start

        #region GetInvoiceHeader
        /// <summary>
        /// GetInvoiceHeader
        /// </summary>
        /// <returns>Datatable Containing GetAllStockValuesLCPVAT</returns>
        public DataTable GetInvoiceHeader()
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.DocNo = DocNo;
                dtTest = objStock.GetInvoiceHeader();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetInvoiceHeader


        #region GetInvoiceDetails
        /// <summary>
        /// GetInvoiceHeader
        /// </summary>
        /// <returns>Datatable Containing GetAllStockValuesLCPVAT</returns>
        public DataTable GetInvoiceDetails()
        {
            DataTable dtTest = null;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.DocNo = DocNo;
                dtTest = objStock.GetInvoiceDetails();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetInvoiceDetails


        #region GetMISReports
        /// <summary>
        /// GetMISReports
        /// </summary>
        /// <returns>Datatable Containing GetMISReports</returns>
        public DataTable GetMISReports()
        {
            // DataTable dtTest = null;
            DataTable dtTest = new DataTable();
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.IntType = IntType;
                objStock.Country = Country;
                objStock.Location = Location;
                objStock.FromDate = FromDate;

                objStock.ToDate = ToDate;
                objStock.LineCode7 = LineCode7;
                objStock.DivisionCode = DivisionCode;
                dtTest = objStock.GetMISReports();

            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetMISReports

        #region GetDCReports
        /// <summary>
        /// GetDCReports
        /// </summary>
        /// <returns>Datatable Containing GetDCReports</returns>
        public DataTable GetDCReports()
        {
            // DataTable dtTest = null;
            DataTable dtTest = new DataTable();
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.IntType = IntType;
                objStock.PackID = PackID;
                objStock.CevaIssueNo = CevaIssueNo;
                objStock.FromDate = FromDate;

                objStock.ToDate = ToDate;
                objStock.PackBarcode = PackBarcode;
                objStock.PONumber = PONumber;
                objStock.LineCode7 = LineCode7;
                dtTest = objStock.GetDCReports();

            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetDCReports

        #region GetSalesData
        /// <summary>
        /// GetSalesData
        /// </summary>
        /// <returns>Datatable Containing GetDCReports</returns>
        public DataTable GetSalesData()
        {
            // DataTable dtTest = null;
            DataTable dtTest = new DataTable();
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.AsOfDate = AsOfDate;
                objStock.Location = Location;
                dtTest = objStock.GetSalesData();

            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetSalesData


        #region GetStockData
        /// <summary>
        /// GetSalesData
        /// </summary>
        /// <returns>Datatable Containing GetDCReports</returns>
        public DataTable GetStockData()
        {
            // DataTable dtTest = null;
            DataTable dtTest = new DataTable();
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.Location = Location;
                dtTest = objStock.GetStockData();

            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetStockData


        #region UpdateUnitPrice
        /// <summary>
        ///UpdateUnitPrice
        /// </summary>
        /// <returns>Result</returns>
        public bool UpdateUnitPrice()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.UnitPrice = UnitPrice;
                objStock.ItemNo = ItemNo;
                objStock.Country = Country;
                Result = objStock.UpdateUnitPrice();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion UpdateUnitPrice


        #region UpdateSalesLine
        /// <summary>
        ///UpdateSalesLine
        /// </summary>
        /// <returns>Result</returns>
        public bool UpdateSalesLine()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.UnitCost = UnitCost;
                objStock.UnitPrice = UnitPrice;
                objStock.ItemNo = ItemNo;
                objStock.Country = Country;

                objStock.LineAmount = LineAmount;
                objStock.DocNo = DocNo;
                Result = objStock.UpdateSalesLine();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion UpdateSalesLine


        //Invoice Report End

        #region GetFinanceReports
        /// <summary>
        /// GetFinanceReports
        /// </summary>
        /// <returns>Datatable Containing GetFinanceReports</returns>
        public DataTable GetFinanceReports()
        {
            // DataTable dtTest = null;
            DataTable dtTest = new DataTable();
            try
            {
                StockDetails objStock = new StockDetails();

                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.Country = Country;
                objStock.ReportType= ReportType;

                dtTest = objStock.GetFinanceReports();

            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetFinanceReports

        #region GetMarkdown
        /// <summary>
        /// GetMarkdown
        /// </summary>
        /// <returns>Datatable Containing GetMarkdown</returns>
        public DataTable GetMarkdown()
        {
            // DataTable dtTest = null;
            DataTable dtTest = new DataTable();
            try
            {
                StockDetails objStock = new StockDetails();
                objStock.ReportType = ReportType;
                objStock.Location = Location;
                dtTest = objStock.GetMarkdown();

            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetFinanceReports

        #region InsertMarkDown
        /// <summary>
        ///InsertMarkDown
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertMarkDown()
        {
            bool Result = false;
            try
            {
                StockDetails objStock = new StockDetails();
                Result = objStock.InsertMarkDown();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertMarkDown
    }
}