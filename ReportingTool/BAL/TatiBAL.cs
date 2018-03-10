#region NameSpace
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using Test.DAL;
using ReportingTool.DAL;
#endregion NameSpace

namespace ReportingTool.BAL
{
    public class TatiBAL
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
        /// Month
        /// </summary>
        public String Month
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
        /// Brand
        /// </summary>
        public string Brand
        {
            get;
            set;
        }
        /// <summary>
        /// ReceiptNo
        /// </summary>
        public string ReceiptNo
        {
            get;
            set;
        }

        /// <summary>
        /// PONumber
        /// </summary>
        public string PONumber
        {
            get;
            set;
        }

        /// <summary>
        /// ContainerNo
        /// </summary>
        public string ContainerNo
        {
            get;
            set;
        }
        
        /// <summary>
        /// GrnNo
        /// </summary>
        public string GrnNo
        {
            get;
            set;
        }

        /// <summary>
        /// FileName
        /// </summary>
        public string FileName
        {
            get;
            set;
        }

        /// <summary>
        /// AllocationNo
        /// </summary>
        public string AllocationNo
        {
            get;
            set;
        }
        /// <summary>
        /// SONumber
        /// </summary>
        public string SONumber
        {
            get;
            set;
        }

        /// <summary>
        /// IssueNo
        /// </summary>
        public string IssueNo
        {
            get;
            set;
        }

        /// <summary>
        /// SellThrough
        /// </summary>
        public decimal SellThrough
        {
            get;
            set;
        }

        /// <summary>
        /// PackBarcode
        /// </summary>
        public string PackBarcode
        {
            get;
            set;
        }

        /// <summary>
        /// GrnDate
        /// </summary>
        public DateTime GrnDate
        {
            get;
            set;
        }

        /// <summary>
        /// GrnLineNo
        /// </summary>
        public int GrnLineNo
        {
            get;
            set;
        }
        /// <summary>
        /// POLineNo
        /// </summary>
        public int POLineNo
        {
            get;
            set;
        }

        /// <summary>
        /// PackId
        /// </summary>
        public string PackId
        {
            get;
            set;
        }
        /// <summary>
        /// PackType
        /// </summary>
        public string PackType
        {
            get;
            set;
        }
        /// <summary>
        /// GrnQty
        /// </summary>
        public decimal GrnQty
        {
            get;
            set;
        }

        /// <summary>
        /// PackOuter
        /// </summary>
        public decimal PackOuter
        {
            get;
            set;
        }
        /// <summary>
        /// Linecode12
        /// </summary>
        public string LineCode12
        {
            get;
            set;
        }
        /// <summary>
        /// Ratio
        /// </summary>
        public decimal Ratio
        {
            get;
            set;
        }
        /// <summary>
        /// ALLSizesInPack
        /// </summary>
        public string AllSizesInPack
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
        /// LinkedPackId
        /// </summary>
        public string LinkedPackId
        {
            get;
            set;
        }

        /// <summary>
        /// DocNo
        /// </summary>
        public string DocNo
        {
            get;
            set;
        }

        /// <summary>
        /// AdjustmentNo
        /// </summary>
        public string AdjustmentNo
        {
            get;
            set;
        }

        /// <summary>
        /// Quantity
        /// </summary>
        public decimal Quantity
        {
            get;
            set;
        }

        /// <summary>
        /// Id
        /// </summary>
        public int Id
        {
            get;
            set;
        }

        /// <summary>
        /// Remarks
        /// </summary>
        public String Remarks
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
                TatiDAL objStock = new TatiDAL();
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


        #region GetAllStockValuesMY
        /// <summary>
        /// Get All Stock Values
        /// </summary>
        /// <returns>Datatable Containing All Stock Values</returns>
        public DataTable GetAllStockValuesMY(string locationCode)
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                //objStock.FromDate = fromDate;
                //objStock.ToDate = toDate;
                objStock.Location = locationCode;
                dtTest = objStock.GetStockValuesMY();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetAllStockValuesMY


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
                TatiDAL objStock = new TatiDAL();

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

        #region GetStockStatusLCPMY
        /// <summary>
        /// Get All Stock Values LCP
        /// </summary>
        /// <returns>Datatable Containing All Stock Values LCP</returns>
        public DataTable GetStockStatusLCPMY(string locationCode)
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.Location = locationCode;
                dtTest = objStock.GetStockStatusLCPMY();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetStockStatusLCPMY


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
                TatiDAL objStock = new TatiDAL();

                objStock.WeekNo = WeekNo;
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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
                objStock.Location = Location;
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
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();
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
                objStock.InsertStockStatusMY();
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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();

                objStock.UaeOffer = UaeOffer;
                objStock.BahrainOffer = BahrainOffer;

                objStock.OmanOffer = OmanOffer;
                objStock.JordanOffer = JordanOffer;
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();

                objStock.DtSource = DtSource;
                objStock.BultInsert();
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
        #endregion BulkInsert

        #region GetVisitorsWeeklyReport
        /// <summary>
        /// Get Visitors Weekly Report
        /// </summary>
        /// <returns>Datatable Containing All Visitors Weekly Report</returns>
        public DataTable GetVisitorsWeeklyReport(string locationCode, DateTime PostingDate)
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objVisitors = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();
                objStock.WeekNo = WeekNo;
                objStock.Year = Year;

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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();

                objStock.Location = Location;
                objStock.CompanyName = CompanyName;
                objStock.AsOfDate = AsOfDate;
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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
                objStock.Country = Country;
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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
                objStock.DtSource = DtSource;
                Result = objStock.ImportSalesPlan();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
                objStock.FromDate = FromDate;
                objStock.Location = Location;

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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
                objStock.PostingDate = PostingDate;
                objStock.ReportType = ReportType;
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
                objStock.Country = Country;
                dtReport = objStock.GetStoreByCountry();
            }
            catch (Exception e)
            {

            }
            return dtReport;
        }
        #endregion Get Store By Country



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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objVisitors = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();

                objStock.ItemOperationType = ItemOperationType;
                objStock.ILEOperationType = ILEOperationType;
                objStock.ValueOperationType = ValueOperationType;

                objStock.FootFallOperationType = FootFallOperationType;
                objStock.TransactionOperationType = TransactionOperationType;

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


        #region GetTatiSalesFile
        /// <summary>
        /// GetTatiSalesFile
        /// </summary>
        /// <returns>Datatable Containing All GetTatiSalesFile</returns>
        public DataTable GetTatiSalesFile()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.Location = Location;
                objStock.PostingDate = PostingDate;
                dtTest = objStock.GetTatiSalesFile();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetTatiSalesFile


        #region GetTatiStockFile
        /// <summary>
        /// GetTatiStockFile
        /// </summary>
        /// <returns>Datatable Containing All GetTatiStockFile</returns>
        public DataTable GetTatiStockFile()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.PostingDate = PostingDate;
                objStock.Location = Location;

                dtTest = objStock.GetTatiStockFile();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetTatiStockFile



        #region InsertCashFlow
        /// <summary>
        /// Insert Visitor Data Tati
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertCashFlow()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.WeekNo = WeekNo;
                objStock.Year=Year;
                Result=objStock.InsertCashFlow();
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
        #endregion InsertCashFlow


        #region InsertCashFlowMY
        /// <summary>
        /// InsertCashFlowMY
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertCashFlowMY()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.WeekNo = WeekNo;
                objStock.Year = Year;
                Result = objStock.InsertCashFlowMY();
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
        #endregion InsertCashFlowMY

        #region GetCashFlowTati
        /// <summary>
        /// GetCashFlowTati
        /// </summary>
        /// <returns>Datatable Containing All GetCashFlowTati</returns>
        public DataTable GetCashFlowTati()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                // objStock.PostingDate = PostingDate;
                objStock.JorRate = JorRate;
                objStock.Location = Location;

                dtTest = objStock.GetCashFlowTati();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetCashFlowTati


        #region GetCashFlowMY
        /// <summary>
        /// GetCashFlowMY
        /// </summary>
        /// <returns>Datatable Containing All GetCashFlowMY</returns>
        public DataTable GetCashFlowMY()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                // objStock.PostingDate = PostingDate;
                objStock.JorRate = JorRate;
                objStock.Location = Location;

                dtTest = objStock.GetCashFlowMY();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetCashFlowMY


        #region GetMonth
        /// <summary>
        /// GetMonth
        /// </summary>
        /// <returns>Datatable Containing All GetMonth</returns>
        public DataTable GetMonth()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                // objStock.PostingDate = PostingDate;

                objStock.Month = Month;
                objStock.Year = Year;

                dtTest = objStock.GetMonth();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetMonth



        #region GetCashFlowBankOpening
        /// <summary>
        /// GetCashFlowBankOpening
        /// </summary>
        /// <returns>Datatable Containing All GetCashFlowBankOpening</returns>
        public DataTable GetCashFlowBankOpening()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.Year = Year;
                objStock.Location = Location;
                objStock.JorRate = JorRate;
                dtTest = objStock.GetCashFlowBankOpening();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetCashFlowBankOpening


        #region GetCashFlowBankOpeningMY
        /// <summary>
        /// GetCashFlowBankOpeningMY
        /// </summary>
        /// <returns>Datatable Containing All GetCashFlowBankOpeningMY</returns>
        public DataTable GetCashFlowBankOpeningMY()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.Year = Year;
                objStock.Location = Location;
                objStock.JorRate = JorRate;
                dtTest = objStock.GetCashFlowBankOpeningMY();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetCashFlowBankOpeningMY



        #region GetProfitAndLossTati
        /// <summary>
        /// GetProfitAndLossTati
        /// </summary>
        /// <returns>Datatable Containing All ProfitAndLossTati</returns>
        public DataTable GetProfitAndLossTati()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.JorRate = JorRate;
                objStock.Location = Location;
                dtTest = objStock.GetProfitLossTati();
                
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetProfitAndLossTati


        #region GetProfitAndLossMY
        /// <summary>
        /// GetProfitAndLossMY
        /// </summary>
        /// <returns>Datatable Containing All GetProfitAndLossMY</returns>
        public DataTable GetProfitAndLossMY()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.JorRate = JorRate;
                objStock.Location = Location;
                dtTest = objStock.GetProfitLossMY();

                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetProfitAndLossMY


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
                TatiDAL objStock = new TatiDAL();

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

        #region InsertProfitAndLossReportMY
        /// <summary>
        /// InsertProfitAndLossReportMY
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertProfitAndLossReportMY()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.MonthNo = MonthNo;
                objStock.Year = Year;
                Result = objStock.InsertProfitLossReportMY();
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
        #endregion InsertProfitAndLossReportMY


        #region GetShipmentReport
        /// <summary>
        /// GetShipmentReport
        /// </summary>
        /// <returns>Datatable Containing All GetShipmentReport</returns>
        public DataTable GetShipmentReport()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                
                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.Location = Location;
                dtTest = objStock.GetShipmentReport();

                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetShipmentReport


        #region GetBrandDetails
        /// <summary>
        /// GetBrandDetails
        /// </summary>
        /// <returns>Datatable Containing All BrandDetails</returns>
        public DataTable GetBrandDetails()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                dtTest = objStock.GetBrandDetails();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetBrandDetails

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
                TatiDAL objStock = new TatiDAL();
                objStock.Type = Type;
                objStock.Brand = Brand;
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
                TatiDAL objStock = new TatiDAL();
                objStock.Location = Location;
                objStock.Brand = Brand;
                Result= objStock.InsertGLAccountDetails();
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();

                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
                objStock.Month = Month;
                objStock.Location = Location;
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
                TatiDAL objStock = new TatiDAL();
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
                TatiDAL objStock = new TatiDAL();

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
                TatiDAL objStock = new TatiDAL();
                objStock.Brand = Brand;
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


        #region InsertProfitLossReportConsolidated
        /// <summary>
        /// InsertProfitLossReportConsolidated
        /// </summary>
        /// <returns> InsertProfitLossReportConsolidated</returns>
        public bool InsertProfitLossReportConsolidated()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.OmanRate = OmanRate;
                objStock.JorRate = JorRate;
                objStock.KsaRate = KsaRate;
                objStock.BahRate = BahRate;

                objStock.UaeRate = UaeRate;

                Result = objStock.InsertProfitLossReportConsolidated();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertProfitLossReportConsolidated

        #region GetProfitLossReportConsolidated
        /// <summary>
        /// GetProfitLossReportConsolidated
        /// </summary>
        /// <returns>Datatable Containing All GetProfitLossReportConsolidated</returns>
        public DataTable GetProfitLossReportConsolidated()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.Brand= Brand;
                dtTest = objStock.GetProfitLossReportConsolidated();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetProfitLossReportConsolidated


        #region GetMYMISReports
        /// <summary>
        /// GetMYMISReports
        /// </summary>
        /// <returns>Datatable Containing All GetMYMISReports</returns>
        public DataTable GetMYMISReports()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.Location =Location;
                objStock.ReportType = ReportType;
                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;
              

                dtTest = objStock.GetMYMISReports();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetMYMISReports


        #region GetTATIMISReports
        /// <summary>
        /// GetTATIMISReports
        /// </summary>
        /// <returns>Datatable Containing All GetMYMISReports</returns>
        public DataTable GetTATIMISReports()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.Location = Location;
                objStock.ReportType = ReportType;
                objStock.FromDate = FromDate;
                objStock.ToDate = ToDate;

                objStock.JordanOffer = JordanOffer;
                objStock.Country = Country;
                dtTest = objStock.GetTATIMISReports();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetTATIMISReports



        //TATI Ends


        #region GetDCAllocation
        /// <summary>
        /// GetDCAllocation
        /// </summary>
        /// <returns>Datatable Containing All GetDCAllocation</returns>
        public DataTable GetDCAllocation()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.SellThrough = SellThrough;
                objStock.Location = Location;
                dtTest = objStock.GetDCAllocation();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetDCAllocation

        #region GetDCAllocationSingle
        /// <summary>
        /// GetDCAllocationSingle
        /// </summary>
        /// <returns>Datatable Containing All GetDCAllocation</returns>
        public DataTable GetDCAllocationSingle()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.Location = Location;
                dtTest = objStock.GetDCAllocationSingle();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetDCAllocationSingle

        #region GetDCStockProcess
        /// <summary>
        /// GetDCStockProcess
        /// </summary>
        /// <returns>Datatable Containing All GetDCStockProcess</returns>
        public DataTable GetDCStockProcess()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                dtTest = objStock.GetDCStockProcess();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetDCStockProcess

        #region GetDCStockSingle
        /// <summary>
        /// GetDCStockSingle
        /// </summary>
        /// <returns>Datatable Containing All GetDCStockSingle</returns>
        public DataTable GetDCStockSingle()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                dtTest = objStock.GetDCStockSingle();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetDCStockSingle


        #region GetMYSalesFile
        /// <summary>
        /// GetMYSalesFile
        /// </summary>
        /// <returns>Datatable Containing All GetMYSalesFile</returns>
        public DataTable GetMYSalesFile()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.Location = Location;
                objStock.PostingDate = PostingDate;
                objStock.ReportType = ReportType;
                objStock.ReceiptNo = ReceiptNo;
                dtTest = objStock.GetMYSalesFile();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetMYSalesFile


        #region ImportPackExtract
        /// <summary>
        /// ImportPackExtract
        /// </summary>
        /// <returns>Result</returns>
        public bool ImportPackExtract()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.DtSource = DtSource;
                objStock.ImportPackExtract();
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
        #endregion ImportPackExtract

        #region ImportContainerExtract
        /// <summary>
        /// ImportContainerExtract
        /// </summary>
        /// <returns>Result</returns>
        public bool ImportContainerExtract()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.ContainerNo = ContainerNo;
                objStock.DtSource = DtSource;
                objStock.ImportContainerExtract();
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
        #endregion ImportContainerExtract

        #region GetPODetails
        /// <summary>
        /// GetPODetails
        /// </summary>
        /// <returns>Datatable Containing All GetPODetails</returns>
        public DataTable GetPODetails()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.ReportType = ReportType;
                objStock.PONumber = PONumber;

                objStock.SONumber = SONumber;
                objStock.AllocationNo = AllocationNo;
                objStock.Location = Location;
                objStock.PackBarcode = PackBarcode;
                objStock.LineCode = LineCode;
                objStock.IssueNo = IssueNo;
                objStock.DocNo = DocNo;
                dtTest = objStock.GetPODetails();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetPODetails

        #region InsertPODetail
        /// <summary>
        /// InsertPODetail
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertPODetail()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.PONumber = PONumber;
                objStock.ContainerNo = ContainerNo;
                objStock.InsertPODetail();
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
        #endregion InsertPODetail

        #region ImportPOGRN
        /// <summary>
        /// ImportPOGRN
        /// </summary>
        /// <returns>Result</returns>
        public bool ImportPOGRN()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.GrnNo =GrnNo;
                objStock.DtSource = DtSource;
                objStock.ImportPOGRN();
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
        #endregion ImportPOGRN

        #region ImportSOIssueNote
        /// <summary>
        /// ImportSOIssueNote
        /// </summary>
        /// <returns>Result</returns>
        public bool ImportSOIssueNote()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.IssueNo = IssueNo;
                objStock.DtSource = DtSource;
                objStock.ImportSOIssueNote();
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
        #endregion ImportSOIssueNote


        #region InsertPOGRNHeader
        /// <summary>
        /// InsertPOGRNHeader
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertPOGRNHeader()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.GrnNo = GrnNo;
                objStock.PONumber = PONumber;
                objStock.FileName = FileName;

                objStock.InsertPOGRNHeader();
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
        #endregion InsertPOGRNHeader


        #region GetGRNLog
        /// <summary>
        /// GetGRNLog
        /// </summary>
        /// <returns>Datatable Containing All GetGRNLog</returns>
        public DataTable GetGRNLog()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.GrnNo = GrnNo;
                dtTest = objStock.GetGRNLog();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetGRNLog

        #region GetSOIssueNoteLog
        /// <summary>
        /// GetSOIssueNoteLog
        /// </summary>
        /// <returns>Datatable Containing All GetSOIssueNoteLog</returns>
        public DataTable GetSOIssueNoteLog()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.IssueNo = IssueNo;
                dtTest = objStock.GetSOIssueNoteLog();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetSOIssueNoteLog


        #region InsertStockLedger
        /// <summary>
        /// InsertStockLedger
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertStockLedger()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.ReportType = ReportType;
                objStock.GrnNo = GrnNo;
                objStock.IssueNo = IssueNo;
                objStock.InsertStockLedger();

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
        #endregion InsertStockLedger

        #region GetPOHeader
        /// <summary>
        /// GetPOHeader
        /// </summary>
        /// <returns>Datatable Containing All GetPOHeader</returns>
        public DataTable GetPOHeader()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                //objStock.GrnNo = GrnNo;
                dtTest = objStock.GetPOHeader();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetPOHeader

        #region ImportAllocation
        /// <summary>
        /// ImportAllocation
        /// </summary>
        /// <returns>Result</returns>
        public bool ImportAllocation()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.DtSource = DtSource;
                objStock.AllocationNo = AllocationNo;
                objStock.ImportAllocation();
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
        #endregion ImportAllocation

        #region InsertAllocationHeader
        /// <summary>
        /// InsertAllocationHeader
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertAllocationHeader()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.AllocationNo = AllocationNo;
                objStock.FileName = FileName;
                objStock.Location = Location;
                objStock.InsertAllocationHeader();
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
        #endregion InsertAllocationHeader

        #region GetStore
        /// <summary>
        /// GetStore
        /// </summary>
        /// <returns>Datatable Containing All GetStore</returns>
        public DataTable GetStore()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                dtTest = objStock.GetStore();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetStore

        #region GetAllocationHeader
        /// <summary>
        /// GetAllocationHeader
        /// </summary>
        /// <returns>Datatable Containing All GetAllocationHeader</returns>
        public DataTable GetAllocationHeader()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.Location = Location;
                dtTest = objStock.GetAllocationHeader();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetAllocationHeader

        #region InsertIssueNoteHeader
        /// <summary>
        /// InsertIssueNoteHeader
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertIssueNoteHeader()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.IssueNo= IssueNo;
                objStock.SONumber = SONumber;
                objStock.FileName = FileName;

                objStock.InsertIssueNoteHeader();
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
        #endregion InsertIssueNoteHeader

        #region GetSOHeader
        /// <summary>
        /// GetSOHeader
        /// </summary>
        /// <returns>Datatable Containing All GetAllocationHeader</returns>
        public DataTable GetSOHeader()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();

                dtTest = objStock.GetSOHeader();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetSOHeader

        #region ImportProductGroupListing
        /// <summary>
        /// ImportProductGroupListing
        /// </summary>
        /// <returns>Result</returns>
        public bool ImportProductGroupListing()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.Location = Location;
                objStock.DtSource = DtSource;
                objStock.ImportProductGroupListing();
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
        #endregion ImportProductGroupListing

        #region ImportFamilyListing
        /// <summary>
        /// ImportFamilyListing
        /// </summary>
        /// <returns>Result</returns>
        public bool ImportFamilyListing()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.DtSource = DtSource;
                objStock.Location = Location;
                objStock.ImportFamilyListing();

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
        #endregion ImportFamilyListing


        #region InsertPOGRN
        /// <summary>
        /// InsertPOGRN
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertPOGRN()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.GrnNo = GrnNo;
                objStock.GrnDate = GrnDate;
                objStock.GrnLineNo = GrnLineNo;
                objStock.PONumber = PONumber;

                objStock.ContainerNo = ContainerNo;
                objStock.POLineNo = POLineNo;
                objStock.LineCode = LineCode;
                objStock.PackId = PackId;

                objStock.PackBarcode = PackBarcode;
                objStock.PackType = PackType;
                objStock.GrnQty = GrnQty;

                objStock.InsertPOGRN();

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
        #endregion InsertPOGRN

        #region GetAllSOHeader
        /// <summary>
        /// GetAllSOHeader
        /// </summary>
        /// <returns>Datatable Containing All GetAllSOHeader</returns>
        public DataTable GetAllSOHeader()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
               
                

                dtTest = objStock.GetAllSOHeader();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetAllSOHeader

        #region CreateSOFromPO
        /// <summary>
        /// InsertPOGRN
        /// </summary>
        /// <returns>Result</returns>
        public bool CreateSOFromPO()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.GrnNo = GrnNo;
                objStock.SONumber = SONumber;
                objStock.Location = Location;
                objStock.AsOfDate = AsOfDate;

                objStock.CompanyName = CompanyName;


                objStock.CreateSOFromPurchaseOrder();

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
        #endregion CreateSOFromPO


        #region InsertPackExtract
        /// <summary>
        /// InsertPackExtract
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertPackExtract()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.PackBarcode = PackBarcode;
                objStock.LineCode = LineCode;
                objStock.PackId = PackId;
                objStock.PackType = PackType;

                objStock.PackOuter = PackOuter;
                objStock.LineCode12 = LineCode12;
                objStock.Ratio = Ratio;
                objStock.AllSizesInPack = AllSizesInPack;

                objStock.PackLevel = PackLevel;
                objStock.LinkedPackId = LinkedPackId;

                objStock.InsertPackExtract();

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
        #endregion InsertPackExtract

        #region GetAllPOHeader
        /// <summary>
        /// GetAllPOHeader
        /// </summary>
        /// <returns>Datatable Containing All GetAllPOHeader</returns>
        public DataTable GetAllPOHeader()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.ReportType = ReportType;
                objStock.DocNo = DocNo;
                dtTest = objStock.GetAllPOHeader();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetAllPOHeader

        #region DeleteDocs
        /// <summary>
        /// DeleteDocs
        /// </summary>
        /// <returns>Result</returns>
        public bool DeleteDocs()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.DocNo = DocNo;
                objStock.ReportType = ReportType;

                objStock.DeleteDocs();

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
        #endregion DeleteDocs

        #region CheckContainerExtract
        /// <summary>
        /// CheckContainerExtract
        /// </summary>
        /// <returns>Datatable Containing All CheckContainerExtract</returns>
        public DataTable CheckContainerExtract()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.LineCode = LineCode;
                objStock.PackId = PackId;
                objStock.PackType = PackType;
                dtTest = objStock.CheckContainerExtract();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion CheckContainerExtract


        #region ImportExtraContainerExtract
        /// <summary>
        /// ImportExtraContainerExtract
        /// </summary>
        /// <returns>Result</returns>
        public bool ImportExtraContainerExtract()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();
                //objStock.ContainerNo = ContainerNo;
                objStock.DtSource = DtSource;
                objStock.ImportExtraContainerExtract();
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
        #endregion ImportExtraContainerExtract

       
        #region DeletePOGrnHeader
        /// <summary>
        /// DeletePOGrnHeader
        /// </summary>
        /// <returns>Result</returns>
        public bool DeletePOGrnHeader()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.GrnNo = GrnNo;
                objStock.DeletePOGrnHeader();
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
        #endregion DeletePOGrnHeader


        #region InsertExtraContainerExtract
        /// <summary>
        /// InsertExtraContainerExtract
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertExtraContainerExtract()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.ContainerNo = ContainerNo;
                objStock.InsertExtraContainerExtract();
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
        #endregion InsertExtraContainerExtract


        #region ImportAdjustment
        /// <summary>
        /// ImportAdjustment
        /// </summary>
        /// <returns>Result</returns>
        public bool ImportAdjustment()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();
                //objStock.ContainerNo = ContainerNo;
                objStock.DtSource = DtSource;
                objStock.ImportAdjustment();
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
        #endregion ImportAdjustment


        #region InsertAdjustmentHeader
        /// <summary>
        /// InsertAdjustmentHeader
        /// </summary>
        /// <returns>Result</returns>
        public bool InsertAdjustmentHeader()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.AdjustmentNo = AdjustmentNo;
                objStock.DocNo = DocNo;
                objStock.FileName = FileName;

                objStock.InsertAdjustmentHeader();

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
        #endregion InsertAdjustmentHeader


        #region GetCheckAdjustmentDetails
        /// <summary>
        /// GetCheckAdjustmentDetails
        /// </summary>
        /// <returns>Datatable Containing All GetAllPOHeader</returns>
        public DataTable GetCheckAdjustmentDetails()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.AdjustmentNo = AdjustmentNo;
                dtTest = objStock.GetCheckAdjustmentDetails();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetCheckAdjustmentDetails



        #region ImportTransferOrder
        /// <summary>
        /// ImportTransferOrder
        /// </summary>
        /// <returns>Result</returns>
        public bool ImportTransferOrder()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.DtSource = DtSource;
                objStock.ImportTransferOrder();

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
        #endregion ImportTransferOrder


        #region InsertTransferHeader
        /// <summary>
        /// InsertTransferHeader
        /// </summary>
        /// <returns> InsertTransferHeader</returns>
        public bool InsertTransferHeader()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.FileName = FileName;
                objStock.DocNo = DocNo;
                Result = objStock.InsertTransferHeader();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertTransferHeader

        #region UpdateTransferOrder
        /// <summary>
        /// UpdateTransferOrder
        /// </summary>
        /// <returns> UpdateTransferOrder</returns>
        public bool UpdateTransferOrder()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.Location = Location;
                objStock.DocNo = DocNo;
                Result = objStock.UpdateTransferOrder();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion UpdateTransferOrder

        #region UpdateReceivedQty
        /// <summary>
        /// UpdateReceivedQty
        /// </summary>
        /// <returns> UpdateReceivedQty</returns>
        public bool UpdateReceivedQty()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.Quantity = Quantity;
                objStock.DocNo = DocNo;
                objStock.PackBarcode = PackBarcode;
                objStock.FileName = FileName;
                Result = objStock.UpdateReceivedQty();

                //cmdTest.Parameters.Add(new SqlParameter("@DocNo", SqlDbType.VarChar)).Value = DocNo;
                //cmdTest.Parameters.Add(new SqlParameter("@Qty", SqlDbType.Decimal)).Value = Quantity;
                //cmdTest.Parameters.Add(new SqlParameter("@Barcode", SqlDbType.VarChar)).Value = PackBarcode;

                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion UpdateReceivedQty

        #region GetTransferOrder
        /// <summary>
        /// GetTransferOrder
        /// </summary>
        /// <returns>Datatable Containing All GetTransferOrder</returns>
        public DataTable GetTransferOrder()
        {
            DataTable dtTest = null;
            try
            {
                TatiDAL objStock = new TatiDAL();
                objStock.DocNo = DocNo;
                dtTest = objStock.GetTransferOrder();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return dtTest;
        }
        #endregion GetTransferOrder

        #region UpdateTransferAdjustment
        /// <summary>
        /// UpdateTransferAdjustment
        /// </summary>
        /// <returns> UpdateTransferAdjustment</returns>
        public bool UpdateTransferAdjustment()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.Quantity = Quantity;
                objStock.DocNo = DocNo;
                objStock.Id= Id;
                objStock.Remarks = Remarks;
                Result = objStock.UpdateTransferAdjustment();

                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion UpdateTransferAdjustment

        #region InsertInventoryNAV
        /// <summary>
        /// InsertInventoryNAV
        /// </summary>
        /// <returns> InsertInventoryNAV</returns>
        public bool InsertInventoryNAV()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();

                objStock.CompanyName = CompanyName;
                objStock.Location = Location;
                objStock.DocNo = DocNo;
                Result = objStock.InsertInventoryNAV();
                // ExceptionMessage = objRole.ExceptionMessage;
            }
            catch (Exception e)
            {
                //Common.LogException("Role.cs", "BAL/Role.cs/GetAllActiveRoles", e.Message);
            }
            return Result;
        }
        #endregion InsertInventoryNAV

        #region ImportHSCode
        /// <summary>
        /// ImportHSCode
        /// </summary>
        /// <returns>Result</returns>
        public bool ImportHSCode()
        {
            bool Result = false;
            try
            {
                TatiDAL objStock = new TatiDAL();
                //objStock.ContainerNo = ContainerNo;
                objStock.DtSource = DtSource;
                objStock.ImportHSCode();
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
        #endregion ImportHSCode

    }
}