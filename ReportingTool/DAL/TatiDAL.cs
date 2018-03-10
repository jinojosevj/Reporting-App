#region NameSpace
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using Test.DAL;
#endregion NameSpace

namespace ReportingTool.DAL
{
    public class TatiDAL
    {
        #region Public Properties

        /// <summary>
        /// Exception message
        /// </summary>
        public string ExceptionMessage
        {
            get;
            set;
        }

        /// <summary>
        /// Location code
        /// </summary>
        public string Location
        {
            get;
            set;
        }

        /// <summary>
        /// From Date
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
        /// DtSource
        /// </summary>
        public DataTable DtSource
        {
            get;
            set;
        }

        /// <summary>
        /// Week No.
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
        ///  Month
        /// </summary>
        public String Month
        {
            get;
            set;
        }

        /// <summary>
        ///  MonthNo
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


        #region GetStockValues
        /// <summary>
        ///  Get Stock Values
        /// </summary>
        /// <returns>Datatable Containing All StockValues</returns>
        public DataTable GetStockValues()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetStockStatusReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;
                //cmdTest.Parameters.Add(new SqlParameter("@FromDate", SqlDbType.DateTime)).Value = FromDate;
                //cmdTest.Parameters.Add(new SqlParameter("@ToDate", SqlDbType.DateTime)).Value = ToDate;
                cmdTest.Parameters.Add(new SqlParameter("@Location", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetStockValues


        #region GetStockValuesMY
        /// <summary>
        ///  Get Stock Values
        /// </summary>
        /// <returns>Datatable Containing All StockValues</returns>
        public DataTable GetStockValuesMY()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetStockStatusReportMY", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;
                //cmdTest.Parameters.Add(new SqlParameter("@FromDate", SqlDbType.DateTime)).Value = FromDate;
                //cmdTest.Parameters.Add(new SqlParameter("@ToDate", SqlDbType.DateTime)).Value = ToDate;
                cmdTest.Parameters.Add(new SqlParameter("@Location", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetStockValuesMY


        #region GetWSSIReport
        /// <summary>
        ///  Get WSSI Report
        /// </summary>
        /// <returns>Datatable Containing WSSI Report</returns>
        public DataTable GetWSSIReport()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetWSSIReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@SeasonCode", SqlDbType.VarChar)).Value = SeasonCode;
                cmdTest.Parameters.Add(new SqlParameter("@WeekNo", SqlDbType.VarChar)).Value = WeekNo;
                cmdTest.Parameters.Add(new SqlParameter("@Year", SqlDbType.VarChar)).Value = Year;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetWSSI_Report

        #region GetWSSIForcast
        /// <summary>
        ///  Get WSSI ForCast
        /// </summary>
        /// <returns>Datatable Containing WSSI Forcast</returns>
        public DataTable GetWSSIForcast()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetWssiForcast", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@WeekNo", SqlDbType.VarChar)).Value = WeekNo;
                cmdTest.Parameters.Add(new SqlParameter("@Year", SqlDbType.VarChar)).Value = Year;

                cmdTest.Parameters.Add(new SqlParameter("@Type", SqlDbType.VarChar)).Value = Type;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetWSSIForcast


        #region GetVisitorsReport
        /// <summary>
        /// Get Visitors Report
        /// </summary>
        /// <returns>Datatable Containing All Visitors Report</returns>
        public DataTable GetVisitorsReport()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetVisitorReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Location", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetVisitorsReport

        #region GetStockStatusLCP
        /// <summary>
        ///  Get Stock Status LCP
        /// </summary>
        /// <returns>Datatable Containing Stock Status LCP</returns>
        public DataTable GetStockStatusLCP()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetStockStatusLCP", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;
                cmdTest.Parameters.Add(new SqlParameter("@Location", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetStockStatusLCP


        #region GetStockStatusLCPMY
        /// <summary>
        ///  GetStockStatusLCPMY
        /// </summary>
        /// <returns>Datatable Containing Stock Status LCP</returns>
        public DataTable GetStockStatusLCPMY()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetStockStatusLCPMY", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;
                cmdTest.Parameters.Add(new SqlParameter("@Location", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetStockStatusLCPMY


        #region GetStockStatusLCPSummery
        /// <summary>
        ///  Get Stock Status LCP Summery
        /// </summary>
        /// <returns>Datatable Containing Stock Status LCP Summery</returns>
        public DataTable GetStockStatusLCPSummery()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetStockStatusLCPSummery", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;
                cmdTest.Parameters.Add(new SqlParameter("@FromDate", SqlDbType.DateTime)).Value = FromDate;
                cmdTest.Parameters.Add(new SqlParameter("@ToDate", SqlDbType.DateTime)).Value = ToDate;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetStockStatusLCPSummery

        #region GetPgCmpReport
        /// <summary>
        ///  Get Pg Cmp Report
        /// </summary>
        /// <returns>Datatable Containing Get Pg Cmp Report</returns>
        public DataTable GetPgCmpReport()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetProductGroupCmpReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Location", SqlDbType.VarChar)).Value = Location;


                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetPgCmpReport


        #region UpdateTables
        /// <summary>
        /// For Updating Item master,Item Ledger Entry,Value Entry Tables
        /// </summary>
        /// <returns></returns>

        public bool UpdateTables()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("UpdateTables", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Item", SqlDbType.Bit)).Value = ItemOperationType;
                cmdTest.Parameters.Add(new SqlParameter("@ILE", SqlDbType.Int)).Value = ILEOperationType;
                cmdTest.Parameters.Add(new SqlParameter("@ValueEntry", SqlDbType.Int)).Value = ValueOperationType;

                cmdTest.Parameters.Add(new SqlParameter("@FootFall", SqlDbType.Bit)).Value = FootFallOperationType;
                cmdTest.Parameters.Add(new SqlParameter("@TransHeader", SqlDbType.Bit)).Value = TransactionOperationType;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;

        }
        #endregion UpdateTables

        #region InsertStockStatus


        public bool InsertStockStatus()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("InsertStockStatusReportTable", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@FromDate", SqlDbType.DateTime)).Value = FromDate;
                cmdTest.Parameters.Add(new SqlParameter("@ToDate", SqlDbType.DateTime)).Value = ToDate;

                //cmdTest.Parameters.Add(new SqlParameter("@Item", SqlDbType.Bit)).Value =ItemOperationType;
                //cmdTest.Parameters.Add(new SqlParameter("@ILE", SqlDbType.Bit)).Value = ILEOperationType;
                //cmdTest.Parameters.Add(new SqlParameter("@ValueEntry", SqlDbType.Bit)).Value = ValueOperationType;
                cmdTest.Parameters.Add(new SqlParameter("@StockStatus", SqlDbType.Bit)).Value = SSOperationType;
                cmdTest.Parameters.Add(new SqlParameter("@StockStatusWeekly", SqlDbType.Bit)).Value = SSWeeklyOperationType;
                cmdTest.Parameters.Add(new SqlParameter("@StockStatusReport", SqlDbType.Bit)).Value = SSReportOperationType;

                cmdTest.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdTest.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;
                cmdTest.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;
                cmdTest.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;

                cmdTest.Parameters.Add(new SqlParameter("@KSARate", SqlDbType.Decimal)).Value = KsaRate;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }


        #endregion InsertStockStatus

        #region InsertStockStatusMY
        /// <summary>
        /// InsertStockStatusMY
        /// </summary>
        /// <returns></returns>

        public bool InsertStockStatusMY()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("InsertStockStatusReportTableMY", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@FromDate", SqlDbType.DateTime)).Value = FromDate;
                cmdTest.Parameters.Add(new SqlParameter("@ToDate", SqlDbType.DateTime)).Value = ToDate;

                //cmdTest.Parameters.Add(new SqlParameter("@Item", SqlDbType.Bit)).Value =ItemOperationType;
                //cmdTest.Parameters.Add(new SqlParameter("@ILE", SqlDbType.Bit)).Value = ILEOperationType;
                //cmdTest.Parameters.Add(new SqlParameter("@ValueEntry", SqlDbType.Bit)).Value = ValueOperationType;
                cmdTest.Parameters.Add(new SqlParameter("@StockStatus", SqlDbType.Bit)).Value = SSOperationType;
                cmdTest.Parameters.Add(new SqlParameter("@StockStatusWeekly", SqlDbType.Bit)).Value = SSWeeklyOperationType;
                cmdTest.Parameters.Add(new SqlParameter("@StockStatusReport", SqlDbType.Bit)).Value = SSReportOperationType;

                cmdTest.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdTest.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;
                cmdTest.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;
                cmdTest.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;

                cmdTest.Parameters.Add(new SqlParameter("@KSARate", SqlDbType.Decimal)).Value = KsaRate;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }
        #endregion InsertStockStatusMY


        #region InsertVisitorsReport
        /// <summary>
        /// Insert Visitors Report
        /// </summary>
        /// <returns></returns>

        public bool InsertVisitorsReport()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("UpdateVisitorReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@PostingDate", SqlDbType.DateTime)).Value = PostingDate;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }


        #endregion InsertVisitorsReport

        #region BultInsert
        /// <summary>
        /// Bult Insert
        /// </summary>
        /// <param name="dtSource"></param>
        public void BultInsert()
        {

            DateTime postingDate = Convert.ToDateTime(DtSource.Rows[0]["Date"]);

            bool Result = DeleteStoreFootFall(postingDate);

            try
            {
                string dbCntStr = (null != ConfigurationManager.ConnectionStrings["TestConnectionString"])
                        ? ConfigurationManager.ConnectionStrings["TestConnectionString"].ConnectionString : "";

                System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(dbCntStr);//add connectionstring here

                bcp.DestinationTableName = "StoreFootFallRegister";//give destination table name

                bcp.ColumnMappings.Add("EntryNo", "EntryNo");//Map all columns

                bcp.ColumnMappings.Add("Date", "Date");

                bcp.ColumnMappings.Add("FromTime", "FromTime");

                bcp.ColumnMappings.Add("ToTime", "ToTime");

                bcp.ColumnMappings.Add("StoreNo", "StoreNo");

                bcp.ColumnMappings.Add("Terminal", "Terminal");

                bcp.ColumnMappings.Add("NoOfIns", "NoOfIns");

                bcp.ColumnMappings.Add("NoOfOuts", "NoOfOuts");

                bcp.ColumnMappings.Add("Entrance", "Entrance");

                // and so on...., maap all source table with your destination table
                if (DtSource.Rows.Count > 0)
                {
                    bcp.WriteToServer(DtSource);
                }
            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
            }
        }

        #endregion BultInsert

        #region GetVisitorsWeeklyReport
        /// <summary>
        /// Get Visitors Weekly Report
        /// </summary>
        /// <returns>Datatable Containing All Visitor's Weekly Report</returns>
        public DataTable GetVisitorsWeeklyReport()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetVisitorsWeeklyReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Location", SqlDbType.VarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@PostingDate", SqlDbType.DateTime)).Value = PostingDate;
                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetVisitorsWeeklyReport

        #region DeleteStoreFootFall


        public bool DeleteStoreFootFall(DateTime PostingDate)
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("DeleteStoreFootFall", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@PostingDate", SqlDbType.DateTime)).Value = PostingDate;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }


        #endregion DeleteStoreFootFall


        #region InsertWssiReport
        /// <summary>
        /// Insert Wssi Report
        /// </summary>
        /// <returns></returns>

        public bool InsertWssiReport()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("Cur_InsertWssiReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@INWeek", SqlDbType.Int)).Value = WeekNo;
                cmdTest.Parameters.Add(new SqlParameter("@INYear", SqlDbType.VarChar)).Value = Year;

                cmdTest.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;
                cmdTest.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;
                cmdTest.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdTest.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;

                cmdTest.Parameters.Add(new SqlParameter("@KsaRate", SqlDbType.Decimal)).Value = KsaRate;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }


        #endregion InsertWssiReport

        #region InsertProductGroupCmpReport
        /// <summary>
        /// Insert Product Group Compare Report
        /// </summary>
        /// <returns></returns>

        public bool InsertProductGroupCmpReport()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("InsertProductGroupCmpReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Weekno", SqlDbType.Int)).Value = WeekNo;
                cmdTest.Parameters.Add(new SqlParameter("@Year", SqlDbType.Int)).Value = IntYear;

                cmdTest.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;
                cmdTest.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;

                cmdTest.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;
                cmdTest.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;

                cmdTest.Parameters.Add(new SqlParameter("@KsaRate", SqlDbType.Decimal)).Value = KsaRate;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }


        #endregion InsertProductGroupCmpReport


        #region UpdateOfferPrice
        /// <summary>
        /// Update Offer Price
        /// </summary>
        /// <returns></returns>

        public bool UpdateOfferPrice()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("UpdateOfferPrice", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@UaeOffer", SqlDbType.NVarChar)).Value = UaeOffer;
                cmdTest.Parameters.Add(new SqlParameter("@BahOffer", SqlDbType.NVarChar)).Value = BahrainOffer;

                cmdTest.Parameters.Add(new SqlParameter("@OmanOffer", SqlDbType.NVarChar)).Value = OmanOffer;
                cmdTest.Parameters.Add(new SqlParameter("@JodOffer", SqlDbType.NVarChar)).Value = JordanOffer;
                cmdTest.Parameters.Add(new SqlParameter("@QarOffer", SqlDbType.NVarChar)).Value = QatarOffer;

                cmdTest.Parameters.Add(new SqlParameter("@KsaOffer", SqlDbType.NVarChar)).Value = KsaOffer;
                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }


        #endregion UpdateOfferPrice

        #region GetProcessStatus
        /// <summary>
        /// Get Process Status
        /// </summary>
        /// <returns>Datatable Containing Process Status</returns>
        public DataTable GetProcessStatus()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetProcessStatus", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Id", SqlDbType.Int)).Value = ProcessStatusId;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetProcessStatus

        #region UpdateProcessStatus
        /// <summary>
        /// Update Process Status
        /// </summary>
        /// <returns></returns>

        public bool UpdateProcessStatus()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("UpdateProcessStatus", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Flag", SqlDbType.Bit)).Value = ProcessStatusFlag;
                cmdTest.Parameters.Add(new SqlParameter("@Id", SqlDbType.Int)).Value = ProcessStatusId;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }


        #endregion UpdateProcessStatus

        #region InsertWssiDivisionReport
        /// <summary>
        /// Insert Wssi Division Report
        /// </summary>
        /// <returns></returns>

        public bool InsertWssiDivisionReport()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("Cur_InsertWssiDivisionReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@INWeek", SqlDbType.Int)).Value = WeekNo;
                cmdTest.Parameters.Add(new SqlParameter("@INYear", SqlDbType.VarChar)).Value = Year;

                cmdTest.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;
                cmdTest.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;
                cmdTest.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdTest.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }


        #endregion InsertWssiDivisionReport

        #region Get WSSI Division Report
        /// <summary>
        /// Get WSSI Division Report
        /// </summary>
        /// <returns>Datatable Containing WSSI Division Report</returns>
        public DataTable GetWSSIDivisionReport()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetWSSIDivisionReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@DivisionCode", SqlDbType.VarChar)).Value = DivisionCode;
                cmdTest.Parameters.Add(new SqlParameter("@WeekNo", SqlDbType.VarChar)).Value = WeekNo;
                cmdTest.Parameters.Add(new SqlParameter("@Year", SqlDbType.VarChar)).Value = Year;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion Get WSSI Division Report


        #region InsertWssiProductGroupReport
        /// <summary>
        /// Insert Wssi Product Group Report
        /// </summary>
        /// <returns></returns>

        public bool InsertWssiProductGroupReport()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("InsertWssiProductGroupReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@FromDate", SqlDbType.DateTime)).Value = FromDate;
                cmdTest.Parameters.Add(new SqlParameter("@ToDate", SqlDbType.DateTime)).Value = ToDate;


                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }


        #endregion InsertWssiProductGroupReport

        #region Get Wssi Product Group Report
        /// <summary>
        /// Get Wssi Product Group Report
        /// </summary>
        /// <returns>Datatable Containing Wssi Product Group Report</returns>
        public DataTable GetWSSIProductGroupReport()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetWssiProductGroupReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion Get Wssi Product Group Report



        //---------------------------

        #region Get Stock Summary
        /// <summary>
        /// Get Stock Summery
        /// </summary>
        /// <returns>Datatable Containing Get Stock Summary</returns>
        public DataTable GetStockSummary()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati(1);

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                if (Location == "0408")
                {
                    cmdTest = new SqlCommand("WML_StockSummary_Dept_DC", cnTest);
                }
                else
                {
                    cmdTest = new SqlCommand("WML_StockSummary_Dept", cnTest);
                }
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@CompanyName", SqlDbType.VarChar)).Value = CompanyName;
                cmdTest.Parameters.Add(new SqlParameter("@LocationCode", SqlDbType.VarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@AsOfDate", SqlDbType.DateTime)).Value = AsOfDate;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion Get Stock Summary

        //---------------------------

        //---------------------------

        #region Get Inventory Report
        /// <summary>
        /// Get Inventory Report
        /// </summary>
        /// <returns>Datatable Containing  Inventory Report</returns>
        public DataTable GetInventoryReport()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;


                cmdTest = new SqlCommand("GenerateInventoryReport", cnTest);

                cmdTest.CommandType = CommandType.StoredProcedure;


                cmdTest.Parameters.Add(new SqlParameter("@FromDate", SqlDbType.DateTime)).Value = FromDate;
                cmdTest.Parameters.Add(new SqlParameter("@ToDate", SqlDbType.DateTime)).Value = ToDate;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion Get Inventory Report

        //---------------------------


        //--Best Seller Report--Start

        #region Insert Best Seller Report
        /// <summary>
        /// Insert Best Seller Report
        /// </summary>
        /// <returns></returns>

        public bool InsertBestSellerReport()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("InsertBestSeller", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@StartDate", SqlDbType.DateTime)).Value = FromDate;
                cmdReport.Parameters.Add(new SqlParameter("@EndDate", SqlDbType.DateTime)).Value = ToDate;
                cmdReport.Parameters.Add(new SqlParameter("@Location", SqlDbType.VarChar)).Value = Location;
                cmdReport.Parameters.Add(new SqlParameter("@Division", SqlDbType.VarChar)).Value = DivisionCode;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }


        #endregion InsertBestSellerReport


        #region Insert Best Seller By Linecode7
        /// <summary>
        /// Insert Best Seller By Linecode7
        /// </summary>
        /// <returns></returns>

        public bool InsertBestSellerByLinecode7()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("InsertBestSellerByLinecode7", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@StartDate", SqlDbType.DateTime)).Value = FromDate;
                cmdReport.Parameters.Add(new SqlParameter("@EndDate", SqlDbType.DateTime)).Value = ToDate;
                cmdReport.Parameters.Add(new SqlParameter("@Location", SqlDbType.VarChar)).Value = Location;
                cmdReport.Parameters.Add(new SqlParameter("@Division", SqlDbType.VarChar)).Value = DivisionCode;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertBestSellerByLinecode7



        #region Get Best Seller Report
        /// <summary>
        /// Get Best Seller Report
        /// </summary>
        /// <returns>Datatable Containing Best Seller Report</returns>
        public DataTable GetBestSellerReport()
        {
            DataTable dtReport = new DataTable();

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnReport = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("GetBestSellerReport", cnReport);

                cmdReport.Parameters.Add(new SqlParameter("@StartDate", SqlDbType.DateTime)).Value = FromDate;
                cmdReport.Parameters.Add(new SqlParameter("@EndDate", SqlDbType.DateTime)).Value = ToDate;
                cmdReport.Parameters.Add(new SqlParameter("@ReportType", SqlDbType.VarChar)).Value = ReportType;
                cmdReport.Parameters.Add(new SqlParameter("@Location", SqlDbType.VarChar)).Value = Location;
                cmdReport.Parameters.Add(new SqlParameter("@DivisionCode", SqlDbType.VarChar)).Value = DivisionCode;


                cmdReport.Parameters.Add(new SqlParameter("@OmanExgRate", SqlDbType.Decimal)).Value = OmanRate;
                cmdReport.Parameters.Add(new SqlParameter("@UaeExgRate", SqlDbType.Decimal)).Value = UaeRate;
                cmdReport.Parameters.Add(new SqlParameter("@JorExgRate", SqlDbType.Decimal)).Value = JorRate;
                cmdReport.Parameters.Add(new SqlParameter("@BahExgRate", SqlDbType.Decimal)).Value = BahRate;

                cmdReport.Parameters.Add(new SqlParameter("@KsaExgRate", SqlDbType.Decimal)).Value = KsaRate;

                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdReport);

                try
                {
                    daStock.Fill(dtReport);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return dtReport;
        }
        #endregion Get Best Seller Report


        #region InsertBestSellerSummeryReport
        /// <summary>
        /// InsertBestSellerSummeryReport
        /// </summary>
        /// <returns></returns>

        public bool InsertBestSellerSummeryReport()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("InsertBestSellerSummeryReport", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@StartDate", SqlDbType.DateTime)).Value = FromDate;
                cmdReport.Parameters.Add(new SqlParameter("@EndDate", SqlDbType.DateTime)).Value = ToDate;
                cmdReport.Parameters.Add(new SqlParameter("@Division", SqlDbType.VarChar)).Value = DivisionCode;

                cmdReport.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;
                cmdReport.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;
                cmdReport.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdReport.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }


        #endregion InsertBestSellerSummeryReport

        #region InsertBestSellerSummeryReportLC7
        /// <summary>
        /// InsertBestSellerSummeryReportLC7
        /// </summary>
        /// <returns></returns>

        public bool InsertBestSellerSummeryReportLC7()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("InsertBestSellerSummeryReportLC7", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@StartDate", SqlDbType.DateTime)).Value = FromDate;
                cmdReport.Parameters.Add(new SqlParameter("@EndDate", SqlDbType.DateTime)).Value = ToDate;
                cmdReport.Parameters.Add(new SqlParameter("@Division", SqlDbType.VarChar)).Value = DivisionCode;

                cmdReport.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;
                cmdReport.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;
                cmdReport.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdReport.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }


        #endregion InsertBestSellerSummeryReportLC7

        #region DeleteBestSellerReport
        /// <summary>
        /// Delete BestSellerReport
        /// </summary>
        public void DeleteBestSellerReport()
        {
            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("truncate table BestSellerReport", cnTest);


                cmdReport.CommandTimeout = 0;

                try
                {
                    cmdReport.ExecuteNonQuery();
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

        }
        #endregion DeleteBestSellerReport
        //--Best Seller Report--End

        //--DCStock  Report--Start

        #region Get DCStock
        /// <summary>
        /// Get DCStock
        /// </summary>
        /// <returns>Get DCStock</returns>
        public DataTable GetDCStock()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetDcStock", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@LocationCode", SqlDbType.VarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@AsOfDate", SqlDbType.DateTime)).Value = AsOfDate;
                cmdTest.Parameters.Add(new SqlParameter("@Type", SqlDbType.Bit)).Value = Type;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion Get DCStock

        //--DCStock  Report--End


        #region Get Visitors Vs Sales
        /// <summary>
        /// Get Visitors Vs Sales
        /// </summary>
        /// <returns>Get Visitors Vs Sales</returns>
        public DataTable GetVisitorsVsSales()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetVisitorsVsSalesReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@PostingDate", SqlDbType.DateTime)).Value = PostingDate;
                cmdTest.Parameters.Add(new SqlParameter("@Location", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion Get Visitors Vs Sales


        // --- Retail KPI Report

        #region Get Retail KPI
        /// <summary>
        /// Get Retail KPI
        /// </summary>
        /// <returns>Get Retail KPI</returns>
        public DataTable GetRetailKPI()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetRetailKPIReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Country", SqlDbType.VarChar)).Value = Country;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion Get Retail KPI


        #region InsertRetailKPI
        /// <summary>
        /// InsertRetailKPI
        /// </summary>
        /// <returns></returns>

        public bool InsertRetailKPI()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("InsertRetailKpiReport", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@WeekNo", SqlDbType.Int)).Value = WeekNo;
                cmdReport.Parameters.Add(new SqlParameter("@Year", SqlDbType.VarChar)).Value = Year;
                cmdReport.Parameters.Add(new SqlParameter("@LYear", SqlDbType.VarChar)).Value = LYear;
                cmdReport.Parameters.Add(new SqlParameter("@L2Year", SqlDbType.VarChar)).Value = L2Year;

                cmdReport.Parameters.Add(new SqlParameter("@fromDate", SqlDbType.DateTime)).Value = FromDate.Date;
                cmdReport.Parameters.Add(new SqlParameter("@toDate", SqlDbType.DateTime)).Value = ToDate.Date;
                cmdReport.Parameters.Add(new SqlParameter("@reportDate", SqlDbType.DateTime)).Value = ReportDate.Date;

                cmdReport.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;
                cmdReport.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;
                cmdReport.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdReport.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;

                cmdReport.Parameters.Add(new SqlParameter("@KsaRate", SqlDbType.Decimal)).Value = KsaRate;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertRetailKPI


        #region InsertRetailKPIByDivision
        /// <summary>
        /// Insert Retail KPI By Division
        /// </summary>
        /// <returns></returns>

        public bool InsertRetailKPIByDivision()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("InsertRetailKPIByDivision", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@WeekNo", SqlDbType.Int)).Value = WeekNo;
                cmdReport.Parameters.Add(new SqlParameter("@Year", SqlDbType.VarChar)).Value = Year;
                cmdReport.Parameters.Add(new SqlParameter("@YearLY", SqlDbType.VarChar)).Value = LYear;
                cmdReport.Parameters.Add(new SqlParameter("@fromDate", SqlDbType.DateTime)).Value = FromDate;
                cmdReport.Parameters.Add(new SqlParameter("@toDate", SqlDbType.DateTime)).Value = ToDate;

                cmdReport.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;
                cmdReport.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;
                cmdReport.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdReport.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;

                cmdReport.Parameters.Add(new SqlParameter("@KsaRate", SqlDbType.Decimal)).Value = KsaRate;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertRetailKPIByDivision


        #region Get Retail KPI By Division
        /// <summary>
        /// Get Retail KPI By Division
        /// </summary>
        /// <returns>Get Retail KPI By Division</returns>
        public DataTable GetRetailKPIByDivision()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetRetailKPIByDivision", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Country", SqlDbType.VarChar)).Value = Country;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion Get Retail KPI By Division


        #region Get Retail KPI LFL
        /// <summary>
        /// Get Retail KPI LFL
        /// </summary>
        /// <returns>Get Retail KPI LFL</returns>
        public DataTable GetRetailKPILFL()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetRetailKpiLFL", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Country", SqlDbType.VarChar)).Value = Country;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion  Get Retail KPI LFL


        #region GetRetailKPIByDivisionLFL
        /// <summary>
        /// GetRetailKPIByDivisionLFL
        /// </summary>
        /// <returns>GetRetailKPIByDivisionLFL</returns>
        public DataTable GetRetailKPIByDivisionLFL()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetRetailKPIByDivisionLFL", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Country", SqlDbType.VarChar)).Value = Country;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion  GetRetailKPIByDivisionLFL

        #region InsertWeeklySales
        /// <summary>
        /// Insert Weekly Sales
        /// </summary>
        /// <returns></returns>

        public bool InsertWeeklySales()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("Cur_InsertWeeklySales", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@FromWeek", SqlDbType.Int)).Value = WeekNo;
                cmdReport.Parameters.Add(new SqlParameter("@ToWeek", SqlDbType.VarChar)).Value = WeekNo;
                cmdReport.Parameters.Add(new SqlParameter("@INYear", SqlDbType.VarChar)).Value = Year;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertWeeklySales

        #region ImportSalesPlan
        /// <summary>
        /// Import Sales Plan
        /// </summary>
        /// <param name="dtSource"></param>
        public bool ImportSalesPlan()
        {

            DateTime postingDate = Convert.ToDateTime(DtSource.Rows[0]["PostingDate"]);

            bool Result = false;

            // DeleteStoreFootFall(postingDate);
            try
            {
                string dbCntStr = (null != ConfigurationManager.ConnectionStrings["TestConnectionString"])
                        ? ConfigurationManager.ConnectionStrings["TestConnectionString"].ConnectionString : "";

                System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(dbCntStr);//add connectionstring here

                bcp.DestinationTableName = "SalesPlan";//give destination table name

                bcp.ColumnMappings.Add("PostingDate", "PostingDate");//Map all columns
                bcp.ColumnMappings.Add("PlanAmount", "PlanAmount");
                bcp.ColumnMappings.Add("StoreCode", "StoreCode");
                bcp.ColumnMappings.Add("WeekNo", "WeekNo");

                // and so on...., maap all source table with your destination table
                if (DtSource.Rows.Count > 0)
                {
                    bcp.WriteToServer(DtSource);
                    Result = true;
                }
            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
                Result = false;
            }

            return Result;
        }

        #endregion ImportSalesPlan


        #region ImportLinearCount
        /// <summary>
        /// Import Linear Count
        /// </summary>
        /// <param name="dtSource"></param>
        public bool ImportLinearCount()
        {
            bool Result = false;

            // DeleteStoreFootFall(postingDate);
            try
            {
                string dbCntStr = (null != ConfigurationManager.ConnectionStrings["TestConnectionString"])
                        ? ConfigurationManager.ConnectionStrings["TestConnectionString"].ConnectionString : "";

                System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(dbCntStr);//add connectionstring here

                bcp.DestinationTableName = "LinearCount";//give destination table name

                bcp.ColumnMappings.Add("CategoryCode", "CategoryCode");//Map all columns
                bcp.ColumnMappings.Add("LinearCount", "LinearCount");
                bcp.ColumnMappings.Add("LocationCode", "LocationCode");
                bcp.ColumnMappings.Add("WeekNo", "WeekNo");

                bcp.ColumnMappings.Add("Year", "Year");
                bcp.ColumnMappings.Add("CreatedDate", "CreatedDate");

                // and so on...., maap all source table with your destination table
                if (DtSource.Rows.Count > 0)
                {
                    bcp.WriteToServer(DtSource);
                    Result = true;
                }
            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
                Result = false;
            }

            return Result;
        }

        #endregion ImportLinearCount


        #region InsertRetailKpiMonth
        /// <summary>
        /// Insert Retail Kpi Month
        /// </summary>
        /// <returns></returns>

        public bool InsertRetailKpiMonth()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("InsertRetailKpiMonth", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@fromDate", SqlDbType.DateTime)).Value = FromDate;
                cmdReport.Parameters.Add(new SqlParameter("@toDate", SqlDbType.DateTime)).Value = ToDate;
                cmdReport.Parameters.Add(new SqlParameter("@fromDateLY", SqlDbType.DateTime)).Value = FromDateLY;
                cmdReport.Parameters.Add(new SqlParameter("@toDateLY", SqlDbType.DateTime)).Value = ToDateLY;

                cmdReport.Parameters.Add(new SqlParameter("@fromDate2LY", SqlDbType.DateTime)).Value = FromDate2LY;
                cmdReport.Parameters.Add(new SqlParameter("@toDate2LY", SqlDbType.DateTime)).Value = ToDate2LY;
                cmdReport.Parameters.Add(new SqlParameter("@fromDateYear", SqlDbType.DateTime)).Value = FromDateYear;
                cmdReport.Parameters.Add(new SqlParameter("@toDateYear", SqlDbType.DateTime)).Value = ToDateYear;

                cmdReport.Parameters.Add(new SqlParameter("@fromDateYearLY", SqlDbType.DateTime)).Value = FromDateYearLY;
                cmdReport.Parameters.Add(new SqlParameter("@toDateYearLY", SqlDbType.DateTime)).Value = ToDateYearLY;
                cmdReport.Parameters.Add(new SqlParameter("@fromDateYear2LY", SqlDbType.DateTime)).Value = FromDateYear2LY;
                cmdReport.Parameters.Add(new SqlParameter("@toDateYear2LY", SqlDbType.DateTime)).Value = ToDateYear2LY;

                cmdReport.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;
                cmdReport.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;
                cmdReport.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdReport.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertRetailKpiMonth

        #region InsertRetailKPIYearByDivision
        /// <summary>
        /// Insert Retail KPI Year By Division
        /// </summary>
        /// <returns></returns>

        public bool InsertRetailKPIYearByDivision()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("InsertRetailKPIYearByDivision", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@fromDate", SqlDbType.DateTime)).Value = FromDateYear;
                cmdReport.Parameters.Add(new SqlParameter("@toDate", SqlDbType.DateTime)).Value = ToDateYear;
                cmdReport.Parameters.Add(new SqlParameter("@fromDateLY", SqlDbType.DateTime)).Value = FromDateYearLY;
                cmdReport.Parameters.Add(new SqlParameter("@toDateLY", SqlDbType.DateTime)).Value = ToDateYearLY;

                cmdReport.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;
                cmdReport.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;
                cmdReport.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdReport.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertRetailKPIYearByDivision

        #region GetRetailKpiYear
        /// <summary>
        /// Get Retail Kpi Year
        /// </summary>
        /// <returns>Get Retail Kpi Year</returns>
        public DataTable GetRetailKpiYear()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetRetailKpiYear", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Country", SqlDbType.VarChar)).Value = Country;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion Get Retail Kpi Year

        #region GetRetailKpiMonth
        /// <summary>
        /// Get Retail Kpi Month
        /// </summary>
        /// <returns>GetRetailKpiMonth</returns>
        public DataTable GetRetailKpiMonth()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetRetailKpiMonth", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Country", SqlDbType.VarChar)).Value = Country;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetRetailKpiMonth

        #region GetRetailKPIYearByDivision
        /// <summary>
        /// GetRetailKPIYearByDivision
        /// </summary>
        /// <returns>GetRetailKPIYearByDivision</returns>
        public DataTable GetRetailKPIYearByDivision()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetRetailKPIYearByDivision", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Country", SqlDbType.VarChar)).Value = Country;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetRetailKPIYearByDivision

        #region DeleteSalesPlan
        /// <summary>
        /// DeleteSalesPlan
        /// </summary>
        /// <returns></returns>

        public bool DeleteSalesPlan()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("DeleteSalesPlan", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@postingDate", SqlDbType.DateTime)).Value = PostingDate;
                cmdReport.Parameters.Add(new SqlParameter("@storeCode", SqlDbType.VarChar)).Value = Location;
                cmdReport.Parameters.Add(new SqlParameter("@weekNo", SqlDbType.Int)).Value = WeekNo;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion DeleteSalesPlan


        #region Delete Linear Count
        /// <summary>
        /// Delete Linear Count
        /// </summary>
        /// <returns></returns>
        public bool DeleteLinearCount()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("DeleteLinearCount", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@locationCode", SqlDbType.VarChar)).Value = Location;
                cmdReport.Parameters.Add(new SqlParameter("@weekNo", SqlDbType.Int)).Value = WeekNo;
                cmdReport.Parameters.Add(new SqlParameter("@year", SqlDbType.VarChar)).Value = Year;
                cmdReport.Parameters.Add(new SqlParameter("@categoryCode", SqlDbType.VarChar)).Value = CategoryCode;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion Delete Linear Count


        #region GetWeekDetails
        /// <summary>
        /// GetWeekDetails
        /// </summary>
        /// <returns>GetWeekDetails</returns>
        public DataTable GetWeekDetails()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetWeekDetails", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@fromDate", SqlDbType.DateTime)).Value = FromDate.Date;
                cmdTest.Parameters.Add(new SqlParameter("@toDate", SqlDbType.DateTime)).Value = ToDate.Date;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetWeekDetails

        #region InsertDailySales
        /// <summary>
        /// Insert Daily Sales
        /// </summary>
        /// <returns></returns>

        public bool InsertDailySales()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("InsertDailySales", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@fromDate", SqlDbType.DateTime)).Value = FromDate.Date;
                cmdReport.Parameters.Add(new SqlParameter("@toDate", SqlDbType.DateTime)).Value = ToDate.Date;
                cmdReport.Parameters.Add(new SqlParameter("@country", SqlDbType.VarChar)).Value = Country;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertDailySales

        // --- Retail KPI Report

        //DSR Report Start

        #region Insert Dsr Report
        /// <summary>
        /// Insert Dsr Report
        /// </summary>
        /// <returns></returns>

        public bool InsertDsrReport()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("InsertDsrReport", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@fromDate", SqlDbType.DateTime)).Value = FromDate.Date;
                cmdReport.Parameters.Add(new SqlParameter("@locationCode", SqlDbType.VarChar)).Value = Location;
                cmdReport.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;
                cmdReport.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;

                cmdReport.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdReport.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;
                cmdReport.Parameters.Add(new SqlParameter("@KsaRate", SqlDbType.Decimal)).Value = KsaRate;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion Insert Dsr Report


        #region GetDsrReport
        /// <summary>
        /// GetDsrReport
        /// </summary>
        /// <returns>GetDsrReport</returns>
        public DataTable GetDsrReport()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetDsrReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@LocationCode", SqlDbType.VarChar)).Value = Location;
                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetDsrReport


        #region Insert Dsr Division
        /// <summary>
        /// Insert Dsr Division
        /// </summary>
        /// <returns></returns>

        public bool InsertDsrDivision()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("InsertDsrDivision", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@fromDate", SqlDbType.DateTime)).Value = FromDate.Date;
                cmdReport.Parameters.Add(new SqlParameter("@toDate", SqlDbType.DateTime)).Value = ToDate.Date;

                cmdReport.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;
                cmdReport.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;

                cmdReport.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdReport.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;
                cmdReport.Parameters.Add(new SqlParameter("@KsaRate", SqlDbType.Decimal)).Value = KsaRate;


                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion Insert Dsr Division


        #region GetDsrDivision
        /// <summary>
        /// GetDsrDivision
        /// </summary>
        /// <returns>GetDsrDivision</returns>
        public DataTable GetDsrDivision()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetDsrDivision", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@LocationCode", SqlDbType.VarChar)).Value = Location;
                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetDsrDivision

        //DSR Report End


        //Customer Count Report Start

        #region Get Customer Count
        /// <summary>
        /// Get Customer Count
        /// </summary>
        /// <returns>GetCustomerCount</returns>

        public DataTable GetCustomerCount()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetCustomerCount", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@fromDate", SqlDbType.DateTime)).Value = FromDate.Date;
                cmdTest.Parameters.Add(new SqlParameter("@toDate", SqlDbType.DateTime)).Value = ToDate.Date;
                cmdTest.Parameters.Add(new SqlParameter("@storeNo", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion Get Customer Count


        //Customer Count Report End



        //Highest Closing Values Report Start

        #region Get Highest Closing Values
        /// <summary>
        /// Highest Closing Values
        /// </summary>
        /// <returns>Highest Closing Values</returns>

        public DataTable GetHighestClosingValues()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetHighestClosingValues", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@PostingDate", SqlDbType.DateTime)).Value = PostingDate.Date;
                cmdTest.Parameters.Add(new SqlParameter("@LocationCode", SqlDbType.VarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@ReportType", SqlDbType.VarChar)).Value = ReportType;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion Get Highest Closing Values


        //Highest Value Repot End


        #region GetPgcmpSummaryByDivision
        /// <summary>
        /// Get Pgcmp Summary By Division
        /// </summary>
        /// <returns>GetPgcmpSummaryByDivision</returns>

        public DataTable GetPgcmpSummaryByDivision()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetPgcmpSummaryByDivision", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                //cmdTest.Parameters.Add(new SqlParameter("@PostingDate", SqlDbType.DateTime)).Value = PostingDate.Date;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetPgcmpSummaryByDivision

        #region Get Store No
        /// <summary>
        /// GetStoreByCountry
        /// </summary>
        /// <returns>GetStoreByCountry</returns>

        public DataTable GetStoreByCountry()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetStoreByCountry", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Country", SqlDbType.VarChar)).Value = Country;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion Get Store By Country

        #region GetItemInfo
        /// <summary>
        /// GetItemInfo
        /// </summary>
        /// <returns>GetItemInfo</returns>

        public DataTable GetItemInfo()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetItemInfo", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Company", SqlDbType.VarChar)).Value = Country;
                cmdTest.Parameters.Add(new SqlParameter("@Location", SqlDbType.VarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@LineCode", SqlDbType.VarChar)).Value = LineCode;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetItemInfo


        //TATI Starts

        #region InsertRetailKPITati
        /// <summary>
        /// InsertRetailKPITati
        /// </summary>
        /// <returns></returns>

        public bool InsertRetailKPITati()
        {
            bool Result = false;

            //TATI DB Connection

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("InsertRetailKpiReport", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@WeekNo", SqlDbType.Int)).Value = WeekNo;
                cmdReport.Parameters.Add(new SqlParameter("@Year", SqlDbType.VarChar)).Value = Year;
                cmdReport.Parameters.Add(new SqlParameter("@LYear", SqlDbType.VarChar)).Value = LYear;
                cmdReport.Parameters.Add(new SqlParameter("@L2Year", SqlDbType.VarChar)).Value = L2Year;

                cmdReport.Parameters.Add(new SqlParameter("@fromDate", SqlDbType.DateTime)).Value = FromDate.Date;
                cmdReport.Parameters.Add(new SqlParameter("@toDate", SqlDbType.DateTime)).Value = ToDate.Date;
                cmdReport.Parameters.Add(new SqlParameter("@reportDate", SqlDbType.DateTime)).Value = ReportDate.Date;

                //cmdReport.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;
                //cmdReport.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;
                cmdReport.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                //cmdReport.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;

                //cmdReport.Parameters.Add(new SqlParameter("@KsaRate", SqlDbType.Decimal)).Value = KsaRate;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertRetailKPITati

        #region Insert Retail KPI By Division Tati
        /// <summary>
        /// Insert Retail KPI By Division Tati
        /// </summary>
        /// <returns></returns>

        public bool InsertRetailKPIByDivisionTati()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("InsertRetailKPIByDivision", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@WeekNo", SqlDbType.Int)).Value = WeekNo;
                cmdReport.Parameters.Add(new SqlParameter("@Year", SqlDbType.VarChar)).Value = Year;
                cmdReport.Parameters.Add(new SqlParameter("@YearLY", SqlDbType.VarChar)).Value = LYear;
                cmdReport.Parameters.Add(new SqlParameter("@fromDate", SqlDbType.DateTime)).Value = FromDate;

                cmdReport.Parameters.Add(new SqlParameter("@toDate", SqlDbType.DateTime)).Value = ToDate;
                //cmdReport.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;
                // cmdReport.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;
                cmdReport.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;

                //cmdReport.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;
                //cmdReport.Parameters.Add(new SqlParameter("@KsaRate", SqlDbType.Decimal)).Value = KsaRate;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion Insert Retail KPI By Division Tati

        #region InsertWeeklySalesTati
        /// <summary>
        /// InsertWeeklySalesTati
        /// </summary>
        /// <returns></returns>
        public bool InsertWeeklySalesTati()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("Cur_InsertWeeklySales", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@FromWeek", SqlDbType.Int)).Value = WeekNo;
                cmdReport.Parameters.Add(new SqlParameter("@ToWeek", SqlDbType.VarChar)).Value = WeekNo;
                cmdReport.Parameters.Add(new SqlParameter("@INYear", SqlDbType.VarChar)).Value = Year;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertWeeklySalesTati


        #region InsertDailySalesTati
        /// <summary>
        /// Insert Daily Sales Tati
        /// </summary>
        /// <returns></returns>

        public bool InsertDailySalesTati()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("InsertDailySales", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@fromDate", SqlDbType.DateTime)).Value = FromDate.Date;
                cmdReport.Parameters.Add(new SqlParameter("@toDate", SqlDbType.DateTime)).Value = ToDate.Date;
                cmdReport.Parameters.Add(new SqlParameter("@country", SqlDbType.VarChar)).Value = Country;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertDailySalesTati

        #region Get Retail KPI Tati
        /// <summary>
        /// Get Retail KPI Tati
        /// </summary>
        /// <returns>Get Retail KPI Tati</returns>
        public DataTable GetRetailKPITati()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetRetailKPIReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Country", SqlDbType.VarChar)).Value = Country;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion Get Retail KPI Tati

        #region Get Retail KPI LFL Tati
        /// <summary>
        /// Get Retail KPI LFL Tati
        /// </summary>
        /// <returns>Get Retail KPI LFL</returns>
        public DataTable GetRetailKPILFLTati()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetRetailKpiLFL", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Country", SqlDbType.VarChar)).Value = Country;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion  Get Retail KPI LFL Tati

        #region Get Retail KPI By Division Tati
        /// <summary>
        /// Get Retail KPI By Division Tati
        /// </summary>
        /// <returns>Get Retail KPI By Division Tati</returns>
        public DataTable GetRetailKPIByDivisionTati()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetRetailKPIByDivision", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Country", SqlDbType.VarChar)).Value = Country;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion Get RetailKPI By Division Tati

        #region Get Retail KPI By Division LFL Tati
        /// <summary>
        /// Get Retail KPI By Division LFL Tati
        /// </summary>
        /// <returns>Get Retail KPI By Division LFL</returns>
        public DataTable GetRetailKPIByDivisionLFLTati()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetRetailKPIByDivisionLFL", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Country", SqlDbType.VarChar)).Value = Country;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion  Get Retail KPI By Division LFL Tati

        #region GetWeekDetailsTati
        /// <summary>
        /// GetWeekDetailsTati
        /// </summary>
        /// <returns>GetWeekDetails</returns>
        public DataTable GetWeekDetailsTati()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetWeekDetails", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@fromDate", SqlDbType.DateTime)).Value = FromDate.Date;
                cmdTest.Parameters.Add(new SqlParameter("@toDate", SqlDbType.DateTime)).Value = ToDate.Date;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetWeekDetailsTati

        #region Insert Dsr Report Tati
        /// <summary>
        /// Insert Dsr Report Tati
        /// </summary>
        /// <returns></returns>
        public bool InsertDsrReportTati()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("InsertDsrReport", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@fromDate", SqlDbType.DateTime)).Value = FromDate.Date;
                cmdReport.Parameters.Add(new SqlParameter("@locationCode", SqlDbType.VarChar)).Value = Location;
                //cmdReport.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;
                // cmdReport.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;

                cmdReport.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                // cmdReport.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;
                //  cmdReport.Parameters.Add(new SqlParameter("@KsaRate", SqlDbType.Decimal)).Value = KsaRate;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion Insert Dsr Report Tati

        #region Insert Dsr Division Tati
        /// <summary>
        /// Insert Dsr Division Tati
        /// </summary>
        /// <returns></returns>
        public bool InsertDsrDivisionTati()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("InsertDsrDivision", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@fromDate", SqlDbType.DateTime)).Value = FromDate.Date;
                cmdReport.Parameters.Add(new SqlParameter("@toDate", SqlDbType.DateTime)).Value = ToDate.Date;

                //cmdReport.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;
                //cmdReport.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;

                cmdReport.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                //cmdReport.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;
                //cmdReport.Parameters.Add(new SqlParameter("@KsaRate", SqlDbType.Decimal)).Value = KsaRate;


                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion Insert Dsr Division Tati

        #region Get Dsr Report Tati
        /// <summary>
        /// Get Dsr Report Tati
        /// </summary>
        /// <returns>Get Dsr Report Tati</returns>
        public DataTable GetDsrReportTati()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetDsrReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@LocationCode", SqlDbType.VarChar)).Value = Location;
                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion Get Dsr Report Tati

        #region Get Dsr Division Tati
        /// <summary>
        /// Get Dsr Division Tati
        /// </summary>
        /// <returns>GetDsrDivisionTati</returns>
        public DataTable GetDsrDivisionTati()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = null;

                cmdTest = new SqlCommand("GetDsrDivision", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@LocationCode", SqlDbType.VarChar)).Value = Location;
                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion Get Dsr Division Tati



        #region DeleteSalesPlanTati
        /// <summary>
        /// DeleteSalesPlanTati
        /// </summary>
        /// <returns></returns>

        public bool DeleteSalesPlanTati()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("DeleteSalesPlan", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@postingDate", SqlDbType.DateTime)).Value = PostingDate;
                cmdReport.Parameters.Add(new SqlParameter("@storeCode", SqlDbType.VarChar)).Value = Location;
                cmdReport.Parameters.Add(new SqlParameter("@weekNo", SqlDbType.Int)).Value = WeekNo;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion DeleteSalesPlanTati

        #region ImportSalesPlanTati
        /// <summary>
        /// ImportSalesPlanTati
        /// </summary>
        /// <param name="dtSource"></param>
        public bool ImportSalesPlanTati()
        {

            DateTime postingDate = Convert.ToDateTime(DtSource.Rows[0]["PostingDate"]);

            bool Result = false;

            // DeleteStoreFootFall(postingDate);
            try
            {
                string dbCntStr = (null != ConfigurationManager.ConnectionStrings["TatiConnectionString"])
                        ? ConfigurationManager.ConnectionStrings["TatiConnectionString"].ConnectionString : "";

                System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(dbCntStr);//add connectionstring here

                bcp.DestinationTableName = "SalesPlan";//give destination table name

                bcp.ColumnMappings.Add("PostingDate", "PostingDate");//Map all columns
                bcp.ColumnMappings.Add("PlanAmount", "PlanAmount");
                bcp.ColumnMappings.Add("StoreCode", "StoreCode");
                bcp.ColumnMappings.Add("WeekNo", "WeekNo");

                // and so on...., maap all source table with your destination table
                if (DtSource.Rows.Count > 0)
                {
                    bcp.WriteToServer(DtSource);
                    Result = true;
                }
            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
                Result = false;
            }

            return Result;
        }

        #endregion ImportSalesPlanTati


        #region Delete Linear Count Tati
        /// <summary>
        /// Delete Linear Count Tati
        /// </summary>
        /// <returns></returns>
        public bool DeleteLinearCountTati()
        {
            bool Result = false;

            DatabaseConnectionTati dbReport = new DatabaseConnectionTati();

            if (dbReport.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbReport.CnDbConnection;

                SqlCommand cmdReport = new SqlCommand("DeleteLinearCount", cnTest);
                cmdReport.CommandType = CommandType.StoredProcedure;

                cmdReport.Parameters.Add(new SqlParameter("@locationCode", SqlDbType.VarChar)).Value = Location;
                cmdReport.Parameters.Add(new SqlParameter("@weekNo", SqlDbType.Int)).Value = WeekNo;
                cmdReport.Parameters.Add(new SqlParameter("@year", SqlDbType.VarChar)).Value = Year;
                cmdReport.Parameters.Add(new SqlParameter("@categoryCode", SqlDbType.VarChar)).Value = CategoryCode;

                cmdReport.CommandTimeout = 0;

                try
                {
                    Result = (cmdReport.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbReport.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbReport.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion Delete Linear Count Tati

        #region ImportLinearCountTati
        /// <summary>
        /// Import Linear Count Tati
        /// </summary>
        /// <param name="dtSource"></param>
        public bool ImportLinearCountTati()
        {
            bool Result = false;

            // DeleteStoreFootFall(postingDate);
            try
            {
                string dbCntStr = (null != ConfigurationManager.ConnectionStrings["TatiConnectionString"])
                        ? ConfigurationManager.ConnectionStrings["TatiConnectionString"].ConnectionString : "";

                System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(dbCntStr);//add connectionstring here

                bcp.DestinationTableName = "LinearCount";//give destination table name

                bcp.ColumnMappings.Add("CategoryCode", "CategoryCode");//Map all columns
                bcp.ColumnMappings.Add("LinearCount", "LinearCount");
                bcp.ColumnMappings.Add("LocationCode", "LocationCode");
                bcp.ColumnMappings.Add("WeekNo", "WeekNo");

                bcp.ColumnMappings.Add("Year", "Year");
                bcp.ColumnMappings.Add("CreatedDate", "CreatedDate");

                // and so on...., maap all source table with your destination table
                if (DtSource.Rows.Count > 0)
                {
                    bcp.WriteToServer(DtSource);
                    Result = true;
                }
            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
                Result = false;
            }

            return Result;
        }

        #endregion ImportLinearCountTati


        #region InsertVisitorsReportTati
        /// <summary>
        /// Insert Visitors Report Tati
        /// </summary>
        /// <returns></returns>

        public bool InsertVisitorsReportTati()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("UpdateVisitorReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@PostingDate", SqlDbType.DateTime)).Value = PostingDate;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }


        #endregion InsertVisitorsReportTati



        #region GetVisitorsReportTati
        /// <summary>
        /// Get Visitors Report Tati
        /// </summary>
        /// <returns>Datatable Containing All VisitorsReportTati</returns>
        public DataTable GetVisitorsReportTati()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetVisitorReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Location", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetVisitorsReportTati


        #region GetVisitorsWeeklyReportTati
        /// <summary>
        /// GetVisitorsWeeklyReportTati
        /// </summary>
        /// <returns>Datatable Containing All VisitorsWeeklyReportTati</returns>
        public DataTable GetVisitorsWeeklyReportTati()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetVisitorsWeeklyReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Location", SqlDbType.VarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@PostingDate", SqlDbType.DateTime)).Value = PostingDate;
                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion VisitorsWeeklyReportTati


        #region InsertVisitorDataTati
        /// <summary>
        /// InsertVisitorDataTati
        /// </summary>
        /// <param name="dtSource"></param>
        public void InsertVisitorDataTati()
        {

            DateTime postingDate = Convert.ToDateTime(DtSource.Rows[0]["Date"]);

            bool Result = DeleteStoreFootFallTati(postingDate);

            try
            {
                string dbCntStr = (null != ConfigurationManager.ConnectionStrings["TatiConnectionString"])
                        ? ConfigurationManager.ConnectionStrings["TatiConnectionString"].ConnectionString : "";

                System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(dbCntStr);//add connectionstring here

                bcp.DestinationTableName = "StoreFootFallRegister";//give destination table name

                bcp.ColumnMappings.Add("EntryNo", "EntryNo");//Map all columns

                bcp.ColumnMappings.Add("Date", "Date");

                bcp.ColumnMappings.Add("FromTime", "FromTime");

                bcp.ColumnMappings.Add("ToTime", "ToTime");

                bcp.ColumnMappings.Add("StoreNo", "StoreNo");

                bcp.ColumnMappings.Add("Terminal", "Terminal");

                bcp.ColumnMappings.Add("NoOfIns", "NoOfIns");

                bcp.ColumnMappings.Add("NoOfOuts", "NoOfOuts");

                bcp.ColumnMappings.Add("Entrance", "Entrance");

                // and so on...., maap all source table with your destination table
                if (DtSource.Rows.Count > 0)
                {
                    bcp.WriteToServer(DtSource);
                }
            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
            }
        }

        #endregion InsertVisitorDataTati


        #region DeleteStoreFootFallTati
        /// <summary>
        /// Delete Store Foot Fall Tati
        /// </summary>
        /// <param name="PostingDate"></param>
        /// <returns></returns>
        public bool DeleteStoreFootFallTati(DateTime PostingDate)
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("DeleteStoreFootFall", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@PostingDate", SqlDbType.DateTime)).Value = PostingDate;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }


        #endregion Delete Store Foot Fall Tati


        #region UpdateTablesTati
        /// <summary>
        /// For Updating Item master,Item Ledger Entry,Value Entry Tables
        /// </summary>
        /// <returns></returns>

        public bool UpdateTablesTati()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("UpdateTables", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Item", SqlDbType.Bit)).Value = ItemOperationType;
                cmdTest.Parameters.Add(new SqlParameter("@ILE", SqlDbType.Int)).Value = ILEOperationType;
                cmdTest.Parameters.Add(new SqlParameter("@ValueEntry", SqlDbType.Int)).Value = ValueOperationType;

                //cmdTest.Parameters.Add(new SqlParameter("@FootFall", SqlDbType.Bit)).Value = FootFallOperationType;
                cmdTest.Parameters.Add(new SqlParameter("@TransHeader", SqlDbType.Bit)).Value = TransactionOperationType;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;

        }
        #endregion UpdateTablesTati

        //TATI Stock Status Start

        #region InsertStockStatusTati
        /// <summary>
        /// InsertStockStatusTati
        /// </summary>
        /// <returns></returns>
        public bool InsertStockStatusTati()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("InsertStockStatusReportTable", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@FromDate", SqlDbType.DateTime)).Value = FromDate;
                cmdTest.Parameters.Add(new SqlParameter("@ToDate", SqlDbType.DateTime)).Value = ToDate;

                //cmdTest.Parameters.Add(new SqlParameter("@Item", SqlDbType.Bit)).Value =ItemOperationType;
                //cmdTest.Parameters.Add(new SqlParameter("@ILE", SqlDbType.Bit)).Value = ILEOperationType;
                //cmdTest.Parameters.Add(new SqlParameter("@ValueEntry", SqlDbType.Bit)).Value = ValueOperationType;
                cmdTest.Parameters.Add(new SqlParameter("@StockStatus", SqlDbType.Bit)).Value = SSOperationType;
                cmdTest.Parameters.Add(new SqlParameter("@StockStatusWeekly", SqlDbType.Bit)).Value = SSWeeklyOperationType;
                cmdTest.Parameters.Add(new SqlParameter("@StockStatusReport", SqlDbType.Bit)).Value = SSReportOperationType;

                cmdTest.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdTest.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;
                cmdTest.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;
                cmdTest.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;

                cmdTest.Parameters.Add(new SqlParameter("@KsaRate", SqlDbType.Decimal)).Value = KsaRate;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }


        #endregion InsertStockStatusTati

        //TATI StockStatus End


        #region GetTatiSalesFile
        /// <summary>
        /// GetTatiSalesFile
        /// </summary>
        /// <returns>Datatable Containing All GetTatiSalesFile</returns>
        public DataTable GetTatiSalesFile()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetSalesFile", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@PostingDate", SqlDbType.Date)).Value = PostingDate;
                cmdTest.Parameters.Add(new SqlParameter("@Location", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
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
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetStockFile", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@PostingDate", SqlDbType.Date)).Value = PostingDate;
                cmdTest.Parameters.Add(new SqlParameter("@Location", SqlDbType.VarChar)).Value = Location;
                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetTatiStockFile

        //TATI Ends

        #region GetCashFlowTati
        /// <summary>
        /// GetCashFlowTati
        /// </summary>
        /// <returns>Datatable Containing All GetCashFlowTati</returns>
        public DataTable GetCashFlowTati()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetCashFlowReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                //cmdTest.Parameters.Add(new SqlParameter("@PostingDate", SqlDbType.Date)).Value = PostingDate;
                cmdTest.Parameters.Add(new SqlParameter("@StoreCode", SqlDbType.VarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.VarChar)).Value = JorRate;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
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
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetCashFlowReportMY", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                //cmdTest.Parameters.Add(new SqlParameter("@PostingDate", SqlDbType.Date)).Value = PostingDate;
                cmdTest.Parameters.Add(new SqlParameter("@StoreCode", SqlDbType.VarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.VarChar)).Value = JorRate;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
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
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetMonth", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Month", SqlDbType.VarChar)).Value = Month;
                cmdTest.Parameters.Add(new SqlParameter("@Year", SqlDbType.VarChar)).Value = Year;
                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetMonth


        #region InsertCashFlow
        /// <summary>
        /// InsertCashFlow
        /// </summary>
        /// <returns></returns>
        public bool InsertCashFlow()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("Cur_InsertCashFlow", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@ToWeek", SqlDbType.Int)).Value = WeekNo;
                cmdTest.Parameters.Add(new SqlParameter("@INYear", SqlDbType.VarChar)).Value = Year;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }


        #endregion InsertCashFlow


        #region InsertCashFlowMY
        /// <summary>
        /// InsertCashFlowMY
        /// </summary>
        /// <returns></returns>
        public bool InsertCashFlowMY()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("Cur_InsertCashFlowMY", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@ToWeek", SqlDbType.Int)).Value = WeekNo;
                cmdTest.Parameters.Add(new SqlParameter("@INYear", SqlDbType.VarChar)).Value = Year;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }


        #endregion InsertCashFlowMY


        #region GetCashFlowBankOpening
        /// <summary>
        /// GetCashFlowBankOpening
        /// </summary>
        /// <returns>Datatable Containing All GetCashFlowBankOpening</returns>
        public DataTable GetCashFlowBankOpening()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetCashFlowBankOpening", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;
               
                cmdTest.Parameters.Add(new SqlParameter("@Year", SqlDbType.VarChar)).Value = Year;
                cmdTest.Parameters.Add(new SqlParameter("@StoreNo", SqlDbType.VarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
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
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetCashFlowBankOpeningMY", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Year", SqlDbType.VarChar)).Value = Year;
                cmdTest.Parameters.Add(new SqlParameter("@StoreNo", SqlDbType.VarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetCashFlowBankOpeningMY


        #region GetProfitLossTati
        /// <summary>
        /// GetProfitLossTati
        /// </summary>
        /// <returns>Datatable Containing All GetProfitLossTati</returns>
        public DataTable GetProfitLossTati()
        {
            DataSet DS = new DataSet();
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetProfitLossReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdTest.Parameters.Add(new SqlParameter("@StoreNo", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(DS);
                    ExceptionMessage = "";
                    dtTest = DS.Tables[0];
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetProfitLossTati


        #region GetProfitLossMY
        /// <summary>
        /// GetProfitLossMY
        /// </summary>
        /// <returns>Datatable Containing All GetProfitLossMY</returns>
        public DataTable GetProfitLossMY()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetProfitLossReportMY", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdTest.Parameters.Add(new SqlParameter("@StoreNo", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetProfitLossMY



        #region InsertProfitLossReport
        /// <summary>
        /// InsertProfitLossReport
        /// </summary>
        /// <returns></returns>
        public bool InsertProfitLossReport()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("Cur_InsertProfitLossReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@ToMonth", SqlDbType.Int)).Value = MonthNo;
                cmdTest.Parameters.Add(new SqlParameter("@INYear", SqlDbType.VarChar)).Value = Year;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertProfitLossReport


        #region InsertProfitLossReportMY
        /// <summary>
        /// InsertProfitLossReportMY
        /// </summary>
        /// <returns></returns>
        public bool InsertProfitLossReportMY()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("Cur_InsertProfitLossReportMY", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@ToMonth", SqlDbType.Int)).Value = MonthNo;
                cmdTest.Parameters.Add(new SqlParameter("@INYear", SqlDbType.VarChar)).Value = Year;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertProfitLossReportMY


        #region GetShipmentReport
        /// <summary>
        /// GetShipmentReport
        /// </summary>
        /// <returns>Datatable Containing All GetShipmentReport</returns>
        public DataTable GetShipmentReport()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetShipmentReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@StoreNo", SqlDbType.VarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@FromDate", SqlDbType.DateTime)).Value = FromDate;
                cmdTest.Parameters.Add(new SqlParameter("@ToDate", SqlDbType.DateTime)).Value = ToDate;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetShipmentReport


      
        //Dynamic Cashflow Start
        
        #region GetBrandDetails
        /// <summary>
        /// Get Brand Details
        /// </summary>
        /// <returns>Datatable Containing All BrandDetails</returns>
        public DataTable GetBrandDetails()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetBrandDetails", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetBrandDetails


        #region GetStoreDetails
        /// <summary>
        /// Get Store Details
        /// </summary>
        /// <returns>Datatable Containing All StoreDetails</returns>
        public DataTable GetStoreDetails()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetStoreDetails", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Type", SqlDbType.Bit)).Value = Type;
                cmdTest.Parameters.Add(new SqlParameter("@Brand", SqlDbType.VarChar)).Value = Brand;
                cmdTest.CommandTimeout = 0;

                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetStoreDetails


        #region InsertGLAccountDetails
        /// <summary>
        /// InsertGLAccountDetails
        /// </summary>
        /// <returns></returns>
        public bool InsertGLAccountDetails()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("InsertGLAccountDetails", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;
                cmdTest.Parameters.Add(new SqlParameter("@StoreNo", SqlDbType.VarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@Brand", SqlDbType.VarChar)).Value = Brand;
                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertGLAccountDetails


        #region GetMonthDetails
        /// <summary>
        /// Get Month Details
        /// </summary>
        /// <returns>Datatable Containing All Month Details</returns>
        public DataTable GetMonthDetails()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetMonthDetails", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@INYear", SqlDbType.VarChar)).Value =Year;
                cmdTest.Parameters.Add(new SqlParameter("@FromMonth", SqlDbType.Int)).Value =FromMonth;
                cmdTest.Parameters.Add(new SqlParameter("@ToMonth", SqlDbType.Int)).Value = ToMonth;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetMonthDetails


        #region UpdateProfitLossActualReport
        /// <summary>
        /// UpdateProfitLossActualReport
        /// </summary>
        /// <returns></returns>
        public bool UpdateProfitLossActualReport()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("UpdateProfitLossActualReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Month", SqlDbType.VarChar)).Value = Month;
                cmdTest.Parameters.Add(new SqlParameter("@FromDate", SqlDbType.DateTime)).Value = FromDate;
                cmdTest.Parameters.Add(new SqlParameter("@ToDate", SqlDbType.DateTime)).Value = ToDate;
                cmdTest.Parameters.Add(new SqlParameter("@StoreNo", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion UpdateProfitLossActualReport


        #region ImportProfitLoseBudgets
        /// <summary>
        /// ImportProfitLoseBudgets
        /// </summary>
        /// <returns></returns>
        public bool ImportProfitLoseBudgets()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("ImportProfitLoseBudgets", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;
                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion ImportProfitLoseBudgets


        #region UpdateProfitLossBudgetReport
        /// <summary>
        /// Update Profit Loss Budget Report
        /// </summary>
        /// <returns></returns>
        public bool UpdateProfitLossBudgetReport()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("UpdateProfitLossBudgetReport", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Month", SqlDbType.VarChar)).Value = Month;
                cmdTest.Parameters.Add(new SqlParameter("@FromDate", SqlDbType.DateTime)).Value = FromDate;
                cmdTest.Parameters.Add(new SqlParameter("@ToDate", SqlDbType.DateTime)).Value = ToDate;
                cmdTest.Parameters.Add(new SqlParameter("@StoreNo", SqlDbType.VarChar)).Value = Location;

                cmdTest.Parameters.Add(new SqlParameter("@Year", SqlDbType.VarChar)).Value = Year;
                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion UpdateProfitLossBudgetReport


        #region DeleteProfitLossReport
        /// <summary>
        /// DeleteProfitLossReport
        /// </summary>
        /// <returns></returns>
        public bool DeleteProfitLossReport()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("delete ProfitLossReport where BrandName='"+Brand+"'", cnTest);
                
                //cmdTest.CommandType = CommandType.StoredProcedure;
                  cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion DeleteProfitLossReport


        #region InsertProfitLossReportConsolidated
        /// <summary>
        /// InsertProfitLossReportConsolidated
        /// </summary>
        /// <returns></returns>
        public bool InsertProfitLossReportConsolidated()
        {
            bool Result = false;

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("UpdateProfitLossReportConsolidated", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@OmanRate", SqlDbType.Decimal)).Value = OmanRate;
                cmdTest.Parameters.Add(new SqlParameter("@JorRate", SqlDbType.Decimal)).Value = JorRate;
                cmdTest.Parameters.Add(new SqlParameter("@UaeRate", SqlDbType.Decimal)).Value = UaeRate;
                cmdTest.Parameters.Add(new SqlParameter("@BahRate", SqlDbType.Decimal)).Value = BahRate;

                cmdTest.Parameters.Add(new SqlParameter("@KsaRate", SqlDbType.Decimal)).Value = KsaRate;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
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
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetProfitLossReportConsolidated", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@Brand", SqlDbType.VarChar)).Value = Brand;
                
                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetProfitLossReportConsolidated

        //Dynamic Cashflow End


        #region GetMYMISReports
        /// <summary>
        /// GetMYMISReports
        /// </summary>
        /// <returns>Datatable Containing All GetMYMISReports</returns>
        public DataTable GetMYMISReports()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetInventoryReportMY", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;
               
                cmdTest.Parameters.Add(new SqlParameter("@StoreNo", SqlDbType.VarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@ReportType", SqlDbType.VarChar)).Value = ReportType;
                cmdTest.Parameters.Add(new SqlParameter("@FromDate", SqlDbType.DateTime)).Value = FromDate;
                cmdTest.Parameters.Add(new SqlParameter("@ToDate", SqlDbType.DateTime)).Value = ToDate;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetMYMISReports

        #region GetTATIMISReports
        /// <summary>
        /// GetTATIMISReports
        /// </summary>
        /// <returns>Datatable Containing All GetTATIMISReports</returns>
        public DataTable GetTATIMISReports()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetTATIMisReports", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@StoreNo", SqlDbType.VarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@ReportType", SqlDbType.VarChar)).Value = ReportType;
                cmdTest.Parameters.Add(new SqlParameter("@FromDate", SqlDbType.Date)).Value = FromDate;
                cmdTest.Parameters.Add(new SqlParameter("@ToDate", SqlDbType.Date)).Value = ToDate;

                cmdTest.Parameters.Add(new SqlParameter("@PromoNos", SqlDbType.NVarChar)).Value = JordanOffer;
                cmdTest.Parameters.Add(new SqlParameter("@Country", SqlDbType.VarChar)).Value = Country;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetTATIMISReports

        #region GetMYSalesFile
        /// <summary>
        /// GetMYSalesFile
        /// </summary>
        /// <returns>Datatable Containing All GetMYSalesFile</returns>
        public DataTable GetMYSalesFile()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionTati dbTest = new DatabaseConnectionTati();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetMYSalesFile", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@ReceiptNo", SqlDbType.VarChar)).Value = ReceiptNo;
                cmdTest.Parameters.Add(new SqlParameter("@PostingDate", SqlDbType.Date)).Value = PostingDate;
                cmdTest.Parameters.Add(new SqlParameter("@Type", SqlDbType.VarChar)).Value = ReportType;
                cmdTest.Parameters.Add(new SqlParameter("@Location", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetMYSalesFile


        #region DCStock APP
        // Code DC application

        #region GetDCAllocation
        /// <summary>
        /// GetDCAllocation
        /// </summary>
        /// <returns>Datatable Containing All GetMYMISReports</returns>
        public DataTable GetDCAllocation()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetDCAllocation", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@StoreCode", SqlDbType.VarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@SellThrough", SqlDbType.Decimal)).Value = SellThrough;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetDCAllocation


        #region GetDCAllocationSingle
        /// <summary>
        /// GetDCAllocationSingle
        /// </summary>
        /// <returns>Datatable Containing All GetMYMISReports</returns>
        public DataTable GetDCAllocationSingle()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetDCAllocationSingle", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@StoreCode", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
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
            DataTable dtTest = new DataTable();

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetDCStock", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

               // cmdTest.Parameters.Add(new SqlParameter("@StoreCode", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
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
            DataTable dtTest = new DataTable();

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetDCStockSingle", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                // cmdTest.Parameters.Add(new SqlParameter("@StoreCode", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetDCStockSingle

       

        #region ImportPackExtract
        /// <summary>
        /// ImportPackExtract
        /// </summary>
        /// <param name="dtSource"></param>
        public void ImportPackExtract()
        {

            try
            {
             
                DatabaseConnectionDC dbConn = new DatabaseConnectionDC();
                System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(dbConn.dbCntStr);//add connectionstring here

                bcp.DestinationTableName = "PackExtract";//give destination table name

                bcp.ColumnMappings.Add("Linecode(7)", "Linecode(7)");//Map all columns
                bcp.ColumnMappings.Add("Pack Id1", "Pack Id");
                bcp.ColumnMappings.Add("Pack Type", "Pack Type");
                bcp.ColumnMappings.Add("Pack Barcode", "Pack Barcode");

                bcp.ColumnMappings.Add("Pack Outer", "Pack Outer");
                bcp.ColumnMappings.Add("Linecode(12)", "Linecode(12)");
                bcp.ColumnMappings.Add("Ratio", "Ratio");
                bcp.ColumnMappings.Add("All Sizes in Pack?", "All Sizes in Pack?");

                bcp.ColumnMappings.Add("Pack Level", "Pack Level");
                bcp.ColumnMappings.Add("Linked Pack ID1", "Linked Pack ID");
                // and so on...., maap all source table with your destination table
                if (DtSource.Rows.Count > 0)
                {
                    bcp.WriteToServer(DtSource);
                }
            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
            }
        }

        #endregion ImportPackExtract

        #region ImportContainerExtract
        /// <summary>
        /// ImportContainerExtract
        /// </summary>
        /// <param name="dtSource"></param>
        public void ImportContainerExtract()
        {

            try
            {

                DatabaseConnectionDC dbConn = new DatabaseConnectionDC();

                SqlConnection cnTest = dbConn.CnDbConnection;
                //SqlCommand cmdTest = new SqlCommand("Delete[ContainerExtract] where [Container Reference] = '"+ContainerNo+"'", cnTest);
                //cmdTest.ExecuteNonQuery();

                System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(dbConn.dbCntStr);//add connectionstring here

                bcp.DestinationTableName = "ContainerExtract";//give destination table name

                bcp.ColumnMappings.Add("Container Reference", "Container Reference");//Map all columns
                bcp.ColumnMappings.Add("Despatch Date", "Despatch Date");
                bcp.ColumnMappings.Add("Sub-Dept", "Sub-Dept");
                bcp.ColumnMappings.Add("Dummy Nest Product", "Dummy Nest Product");

                bcp.ColumnMappings.Add("Linecode", "Linecode");
                bcp.ColumnMappings.Add("Pack ID1", "Pack ID");
                bcp.ColumnMappings.Add("Description", "Description");
                bcp.ColumnMappings.Add("Delivery location", "Delivery location");

                bcp.ColumnMappings.Add("Season", "Season");
                bcp.ColumnMappings.Add("Pack Qty", "Outer");
                bcp.ColumnMappings.Add("Pack Type", "Pack Type");
                bcp.ColumnMappings.Add("Qty", "Qty");

                // and so on...., maap all source table with your destination table
                if (DtSource.Rows.Count > 0)
                {
                    bcp.WriteToServer(DtSource);
                }
            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
            }
        }

        #endregion ImportContainerExtract


        #region GetPODetails
        /// <summary>
        /// GetPODetails
        /// </summary>
        /// <returns>Datatable Containing All GetPODetails</returns>
        public DataTable GetPODetails()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetPODetails", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@PONumber", SqlDbType.VarChar)).Value = PONumber;
                cmdTest.Parameters.Add(new SqlParameter("@Type", SqlDbType.VarChar)).Value = ReportType;
                cmdTest.Parameters.Add(new SqlParameter("@SONumber", SqlDbType.VarChar)).Value = SONumber;
                cmdTest.Parameters.Add(new SqlParameter("@AllocationNo", SqlDbType.VarChar)).Value = AllocationNo;
                cmdTest.Parameters.Add(new SqlParameter("@StoreNo", SqlDbType.VarChar)).Value = Location;

                cmdTest.Parameters.Add(new SqlParameter("@LineCode7", SqlDbType.VarChar)).Value = LineCode;
                cmdTest.Parameters.Add(new SqlParameter("@PackBarcode", SqlDbType.VarChar)).Value = PackBarcode;
                cmdTest.Parameters.Add(new SqlParameter("@IssueNo", SqlDbType.VarChar)).Value = IssueNo;
                cmdTest.Parameters.Add(new SqlParameter("@DocNo", SqlDbType.VarChar)).Value = DocNo;

                cmdTest.Parameters.Add(new SqlParameter("@CreatedId", SqlDbType.VarChar)).Value = "Admin";
                cmdTest.Parameters.Add(new SqlParameter("@Status", SqlDbType.Bit)).Value = false;
                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetPODetails

        #region InsertPODetail
        /// <summary>
        /// InsertPODetail
        /// </summary>
        /// <returns></returns>
        public bool InsertPODetail()
        {
            bool Result = false;

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("InsertPODetails", cnTest);

                cmdTest.Parameters.Add(new SqlParameter("@PONumber", SqlDbType.VarChar)).Value = PONumber;
                cmdTest.Parameters.Add(new SqlParameter("@ContainerNo", SqlDbType.VarChar)).Value = ContainerNo;
                cmdTest.Parameters.Add(new SqlParameter("@CreatedId", SqlDbType.VarChar)).Value = "Admin";

                cmdTest.CommandType = CommandType.StoredProcedure;
                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertPODetail


        #region InsertPOGRNHeader
        /// <summary>
        /// InsertPOGRNHeader
        /// </summary>
        /// <returns></returns>
        public bool InsertPOGRNHeader()
        {
            bool Result = false;

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("InsertPOGRNHeader", cnTest);

                cmdTest.Parameters.Add(new SqlParameter("@GRNNo", SqlDbType.VarChar)).Value = GrnNo;
                cmdTest.Parameters.Add(new SqlParameter("@PONumber", SqlDbType.VarChar)).Value = PONumber;
                cmdTest.Parameters.Add(new SqlParameter("@FileName", SqlDbType.VarChar)).Value = FileName;
                cmdTest.Parameters.Add(new SqlParameter("@CreatedId", SqlDbType.VarChar)).Value = "Admin";

              
                cmdTest.CommandType = CommandType.StoredProcedure;
                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertPOGRNHeader

        #region ImportPOGRN
        /// <summary>
        /// ImportPOGRN
        /// </summary>
        /// <param name="dtSource"></param>
        public void ImportPOGRN()
        {

            try
            {

                DatabaseConnectionDC dbConn = new DatabaseConnectionDC();

                SqlConnection cnTest = dbConn.CnDbConnection;
                //SqlCommand cmdTest = new SqlCommand("Delete[POGRNHeader] where [GRNNo] = '" + GrnNo + "'", cnTest);
                //cmdTest.ExecuteNonQuery();

                System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(dbConn.dbCntStr);//add connectionstring here

                bcp.DestinationTableName = "POGRNDetails";//give destination table name

                bcp.ColumnMappings.Add("PAGRNNO", "GRNNo");//Map all columns
                bcp.ColumnMappings.Add("PAGRNDATE", "GRNDate");
                bcp.ColumnMappings.Add("PAGRNLINENO", "GRNLineNo");
                bcp.ColumnMappings.Add("PONUMBER", "PONumber");

                bcp.ColumnMappings.Add("CONTAINERNO", "ContainerNo");
                bcp.ColumnMappings.Add("POLINENO", "POLineNo");
                bcp.ColumnMappings.Add("LINECODE7", "LineCode7");
                bcp.ColumnMappings.Add("PACKID1", "PackID");

                bcp.ColumnMappings.Add("PACKBARCODE", "PackBarcode");
                bcp.ColumnMappings.Add("PACKTYPE", "PackType");
                bcp.ColumnMappings.Add("PAGRNQTY", "GRNQty");

                //PAGRNNO PAGRNDATE   PAGRNLINENO PONUMBER    CONTAINERNO POLINENO    LINECODE7 PACKID  PACKBARCODE PACKTYPE    PAGRNQTY


                // and so on...., maap all source table with your destination table
                if (DtSource.Rows.Count > 0)
                {
                    bcp.WriteToServer(DtSource);
                }
            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
            }
        }

        #endregion ImportPOGRN


        #region ImportSOIssueNote
        /// <summary>
        /// ImportSOIssueNote
        /// </summary>
        /// <param name="dtSource"></param>
        public void ImportSOIssueNote()
        {

            try
            {

                DatabaseConnectionDC dbConn = new DatabaseConnectionDC();

                SqlConnection cnTest = dbConn.CnDbConnection;
                //SqlCommand cmdTest = new SqlCommand("Delete[POGRNHeader] where [GRNNo] = '" + GrnNo + "'", cnTest);
                //cmdTest.ExecuteNonQuery();

                System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(dbConn.dbCntStr);//add connectionstring here

                bcp.DestinationTableName = "IssueNoteDetails";//give destination table name

                bcp.ColumnMappings.Add("PAISSUENO", "IssueNo");//Map all columns
                bcp.ColumnMappings.Add("PAISSUEDATE1", "IssueDate");
                bcp.ColumnMappings.Add("PAISSUELINENO", "IssueLineNo");
                bcp.ColumnMappings.Add("SONUMBER", "SONumber");

                bcp.ColumnMappings.Add("SOLINENO", "SOLineNo");
                bcp.ColumnMappings.Add("LINECODE7", "LineCode7");
                bcp.ColumnMappings.Add("PACKID1", "PackID");
                bcp.ColumnMappings.Add("PACKBARCODE", "PackBarcode");

                bcp.ColumnMappings.Add("PACKTYPE", "PackType");
                bcp.ColumnMappings.Add("PAISSUEQTY", "IssueQty");

                // and so on...., maap all source table with your destination table
                if (DtSource.Rows.Count > 0)
                {
                    bcp.WriteToServer(DtSource);
                }
                else
                {
                    ExceptionMessage = "No Data";
                }


               
            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
            }
        }

        #endregion ImportSOIssueNote


        #region GetGRNLog
        /// <summary>
        /// GetGRNLog
        /// </summary>
        /// <returns>Datatable Containing All GetGRNLog</returns>
        public DataTable GetGRNLog()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("CheckPOGRN", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@GRNNo", SqlDbType.VarChar)).Value = GrnNo;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
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
            DataTable dtTest = new DataTable();

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("CheckSOIssueNote", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@IssueNo", SqlDbType.VarChar)).Value = IssueNo;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetSOIssueNoteLog


        #region InsertStockLedger
        /// <summary>
        /// InsertStockLedger
        /// </summary>
        /// <returns></returns>
        public bool InsertStockLedger()
        {
            bool Result = false;

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("InsertStockLedger", cnTest);

                cmdTest.Parameters.Add(new SqlParameter("@ReportType", SqlDbType.VarChar)).Value = ReportType;
                cmdTest.Parameters.Add(new SqlParameter("@IssueNo", SqlDbType.VarChar)).Value = IssueNo;
                cmdTest.Parameters.Add(new SqlParameter("@GRNNo", SqlDbType.VarChar)).Value = GrnNo;
                cmdTest.Parameters.Add(new SqlParameter("@CreatedId", SqlDbType.VarChar)).Value = "Admin";

                cmdTest.CommandType = CommandType.StoredProcedure;
                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
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
            DataTable dtTest = new DataTable();

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetPOHeader", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                //cmdTest.Parameters.Add(new SqlParameter("@GRNNo", SqlDbType.VarChar)).Value = GrnNo;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetPOHeader


        #region InsertAllocationHeader
        /// <summary>
        /// InsertAllocationHeader
        /// </summary>
        /// <returns></returns>
        public bool InsertAllocationHeader()
        {
            bool Result = false;

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("InsertAllocationHeader", cnTest);

                cmdTest.Parameters.Add(new SqlParameter("@AllocationNo", SqlDbType.VarChar)).Value = AllocationNo;
                cmdTest.Parameters.Add(new SqlParameter("@FileName", SqlDbType.VarChar)).Value = FileName;
                cmdTest.Parameters.Add(new SqlParameter("@StoreNo", SqlDbType.VarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@CreatedId", SqlDbType.VarChar)).Value = "Admin";

                cmdTest.Parameters.Add(new SqlParameter("@Status", SqlDbType.Bit)).Value = false;


                cmdTest.CommandType = CommandType.StoredProcedure;
                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertAllocationHeader

        #region ImportAllocation
        /// <summary>
        /// ImportAllocation
        /// </summary>
        /// <param name="dtSource"></param>
        public void ImportAllocation()
        {

            try
            {

                DatabaseConnectionDC dbConn = new DatabaseConnectionDC();

                SqlConnection cnTest = dbConn.CnDbConnection;
               
                System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(dbConn.dbCntStr);//add connectionstring here

                bcp.DestinationTableName = "AllocationDetails";//give destination table name

                //ProductGroup LineCode7  PackId Packbarcode Packtype PackQty LineCode7Qty deptcode


                bcp.ColumnMappings.Add("AllocationNo", "AllocationNo");
                bcp.ColumnMappings.Add("StoreNo", "StoreNo");
                bcp.ColumnMappings.Add("ProductGroup", "ProductGroup");//Map all columns
                bcp.ColumnMappings.Add("LineCode7", "Linecode7/12");
                bcp.ColumnMappings.Add("PackId1", "PackId");
                bcp.ColumnMappings.Add("Packbarcode", "PackBarcode");

                bcp.ColumnMappings.Add("Packtype", "PackType");
                bcp.ColumnMappings.Add("PackQty", "PackQty");
                bcp.ColumnMappings.Add("LineCode7Qty", "LineCode7Qty");
                bcp.ColumnMappings.Add("deptcode", "DeptCode");

                // and so on...., maap all source table with your destination table
                if (DtSource.Rows.Count > 0)
                {
                    bcp.WriteToServer(DtSource);
                }
            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
                DeleteAllocationHeader();
            }
        }

        #endregion ImportAllocation


        #region GetStore
        /// <summary>
        /// GetStore
        /// </summary>
        /// <returns>Datatable Containing All GetStore</returns>
        public DataTable GetStore()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetStore", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                //cmdTest.Parameters.Add(new SqlParameter("@GRNNo", SqlDbType.VarChar)).Value = GrnNo;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
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
            DataTable dtTest = new DataTable();

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetAllocationHeader", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@StoreNo", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetAllocationHeader

        #region InsertIssueNoteHeader
        /// <summary>
        /// InsertIssueNoteHeader
        /// </summary>
        /// <returns></returns>
        public bool InsertIssueNoteHeader()
        {
            bool Result = false;

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("InsertIssueNoteHeader", cnTest);

                cmdTest.Parameters.Add(new SqlParameter("@IssueNo", SqlDbType.VarChar)).Value = IssueNo;
                cmdTest.Parameters.Add(new SqlParameter("@SONumber", SqlDbType.VarChar)).Value = SONumber;
                cmdTest.Parameters.Add(new SqlParameter("@FileName", SqlDbType.VarChar)).Value = FileName;
                cmdTest.Parameters.Add(new SqlParameter("@CreatedId", SqlDbType.VarChar)).Value = "Admin";


                cmdTest.CommandType = CommandType.StoredProcedure;
                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertIssueNoteHeader


        #region GetSOHeader
        /// <summary>
        /// GetSOHeader
        /// </summary>
        /// <returns>Datatable Containing All GetSOHeader</returns>
        public DataTable GetSOHeader()
        {
            DataTable dtTest = new DataTable();

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetSOHeader", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                //cmdTest.Parameters.Add(new SqlParameter("@StoreNo", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetSOHeader


        #region ImportProductGroupListing
        /// <summary>
        /// ImportProductGroupListing
        /// </summary>
        /// <param name="dtSource"></param>
        public void ImportProductGroupListing()
        {

            try
            {
                

                DatabaseConnectionDC dbConn = new DatabaseConnectionDC();

                SqlConnection cnTest = dbConn.CnDbConnection;
                
                SqlCommand cmdTest = new SqlCommand("Delete[LineListingDetail] where [StoreNo] = '" + Location + "'", cnTest);
                cmdTest.ExecuteNonQuery();


                System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(dbConn.dbCntStr);//add connectionstring here

                bcp.DestinationTableName = "LineListingDetail";//give destination table name

                bcp.ColumnMappings.Add("StoreNo1", "StoreNo");
                bcp.ColumnMappings.Add("ProductGroup", "ProductGroup");//Map all columns
                bcp.ColumnMappings.Add("Allowed", "Allowed");
                bcp.ColumnMappings.Add("PackQty", "PackQty");
               

                // and so on...., maap all source table with your destination table
                if (DtSource.Rows.Count > 0)
                {
                    bcp.WriteToServer(DtSource);
                }
            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
            }
        }

        #endregion ImportProductGroupListing


        #region ImportFamilyListing
        /// <summary>
        /// ImportFamilyListing
        /// </summary>
        /// <param name="dtSource"></param>
        public void ImportFamilyListing()
        {

            try
            {

                DatabaseConnectionDC dbConn = new DatabaseConnectionDC();

                SqlConnection cnTest = dbConn.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("Delete[SngPGItemFamilyQtyAllocation] where [StoreCode] = '" + Location + "'", cnTest);
                cmdTest.ExecuteNonQuery();


                System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(dbConn.dbCntStr);//add connectionstring here

                bcp.DestinationTableName = "SngPGItemFamilyQtyAllocation";//give destination table name

                bcp.ColumnMappings.Add("ProductGroup", "ProductGroup");
                bcp.ColumnMappings.Add("ItemFamily", "ItemFamily");//Map all columns
                bcp.ColumnMappings.Add("RequiredStoreQty", "RequiredStoreQty");
                bcp.ColumnMappings.Add("StoreCode1", "StoreCode");


                // and so on...., maap all source table with your destination table
                if (DtSource.Rows.Count > 0)
                {
                    bcp.WriteToServer(DtSource);
                }
            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
            }
        }

        #endregion ImportFamilyListing

        #region InsertPOGRN
        /// <summary>
        /// InsertPOGRN
        /// </summary>
        /// <returns></returns>
        public bool InsertPOGRN()
        {
            bool Result = false;

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("InsertPOGRNDetails", cnTest);

                cmdTest.Parameters.Add(new SqlParameter("@GRNNo", SqlDbType.VarChar)).Value = GrnNo;
                cmdTest.Parameters.Add(new SqlParameter("@GRNDate", SqlDbType.Date)).Value = GrnDate;
                cmdTest.Parameters.Add(new SqlParameter("@GRNLineNo", SqlDbType.Int)).Value = GrnLineNo;
                cmdTest.Parameters.Add(new SqlParameter("@PONumber", SqlDbType.VarChar)).Value = PONumber;

                cmdTest.Parameters.Add(new SqlParameter("@ContainerNo", SqlDbType.VarChar)).Value = ContainerNo;
                cmdTest.Parameters.Add(new SqlParameter("@POLineNo", SqlDbType.Int)).Value = POLineNo;
                cmdTest.Parameters.Add(new SqlParameter("@LineCode7", SqlDbType.VarChar)).Value = LineCode;
                cmdTest.Parameters.Add(new SqlParameter("@PackID", SqlDbType.VarChar)).Value = PackId;

                cmdTest.Parameters.Add(new SqlParameter("@PackBarcode", SqlDbType.VarChar)).Value = PackBarcode;
                cmdTest.Parameters.Add(new SqlParameter("@PackType", SqlDbType.VarChar)).Value = PackType;
                cmdTest.Parameters.Add(new SqlParameter("@GRNQty", SqlDbType.Decimal)).Value = GrnQty;

                cmdTest.CommandType = CommandType.StoredProcedure;
                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
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
            DataTable dtTest = new DataTable();

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetAllSOHeader", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                //cmdTest.Parameters.Add(new SqlParameter("@StoreNo", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetAllSOHeader

        #region Create SO From Purchase Order
        /// <summary>
        /// Create SO From Purchase Order
        /// </summary>
        /// <returns></returns>
        public bool CreateSOFromPurchaseOrder()
        {
            bool Result = false;

            DatabaseConnection dbTest = new DatabaseConnection(1);

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("spInsertSOFromGRN", cnTest);

                cmdTest.Parameters.Add(new SqlParameter("@strGRNNumber", SqlDbType.NVarChar)).Value = GrnNo;
                cmdTest.Parameters.Add(new SqlParameter("@strSO", SqlDbType.NVarChar)).Value = SONumber;
                cmdTest.Parameters.Add(new SqlParameter("@storecode", SqlDbType.NVarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@sodate", SqlDbType.DateTime)).Value =AsOfDate;

                cmdTest.Parameters.Add(new SqlParameter("@CompanyDataBaseName", SqlDbType.NVarChar)).Value = CompanyName;
                
                cmdTest.CommandType = CommandType.StoredProcedure;
                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion CreateSOFromPurchaseOrder


        #region InsertPackExtract
        /// <summary>
        /// InsertPackExtract
        /// </summary>
        /// <returns></returns>
        public bool InsertPackExtract()
        {
            bool Result = false;

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("InsertPackExtract", cnTest);

                cmdTest.Parameters.Add(new SqlParameter("@PackBarcode", SqlDbType.NVarChar)).Value = PackBarcode;
                cmdTest.Parameters.Add(new SqlParameter("@Linecode7", SqlDbType.NVarChar)).Value = LineCode;
                cmdTest.Parameters.Add(new SqlParameter("@PackId", SqlDbType.NVarChar)).Value = PackId;
                cmdTest.Parameters.Add(new SqlParameter("@PackType", SqlDbType.NVarChar)).Value = PackType;

                cmdTest.Parameters.Add(new SqlParameter("@PackOuter", SqlDbType.Decimal)).Value = PackOuter;
                cmdTest.Parameters.Add(new SqlParameter("@Linecode12", SqlDbType.NVarChar)).Value = LineCode12;
                cmdTest.Parameters.Add(new SqlParameter("@Ratio", SqlDbType.Decimal)).Value = Ratio;
                cmdTest.Parameters.Add(new SqlParameter("@ALLSizesInPack", SqlDbType.NVarChar)).Value = AllSizesInPack;
                
                cmdTest.Parameters.Add(new SqlParameter("@PackLevel", SqlDbType.NVarChar)).Value = PackLevel;
                cmdTest.Parameters.Add(new SqlParameter("@LinkedPackId", SqlDbType.NVarChar)).Value = LinkedPackId;
               
                cmdTest.CommandType = CommandType.StoredProcedure;
                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
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
            DataTable dtTest = new DataTable();

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetAllPOHeader", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@DocNo", SqlDbType.VarChar)).Value = DocNo;
                cmdTest.Parameters.Add(new SqlParameter("@Type", SqlDbType.VarChar)).Value = ReportType;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetAllPOHeader


        #region DeleteDocs
        /// <summary>
        /// Delete Docs
        /// </summary>
        /// <returns></returns>
        public bool DeleteDocs()
        {
            bool Result = false;

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("DeleteDocs", cnTest);

                cmdTest.Parameters.Add(new SqlParameter("@DocNo", SqlDbType.VarChar)).Value = DocNo;
                cmdTest.Parameters.Add(new SqlParameter("@UserId", SqlDbType.VarChar)).Value = "Admin";
                cmdTest.Parameters.Add(new SqlParameter("@Type", SqlDbType.NVarChar)).Value = ReportType;
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
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
            DataTable dtTest = new DataTable();

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("CheckContainerExtract", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@LineCode", SqlDbType.VarChar)).Value = LineCode;
                cmdTest.Parameters.Add(new SqlParameter("@PackId", SqlDbType.VarChar)).Value = PackId;
                cmdTest.Parameters.Add(new SqlParameter("@PackType", SqlDbType.VarChar)).Value = PackType;
                
                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion CheckContainerExtract


        #region ImportExtraContainerExtract
        /// <summary>
        /// ImportExtraContainerExtract
        /// </summary>
        /// <param name="dtSource"></param>
        public void ImportExtraContainerExtract()
        {

            try
            {

                DatabaseConnectionDC dbConn = new DatabaseConnectionDC();

                SqlConnection cnTest = dbConn.CnDbConnection;
                //SqlCommand cmdTest = new SqlCommand("Delete[ContainerExtract] where [Container Reference] = '"+ContainerNo+"'", cnTest);
                //cmdTest.ExecuteNonQuery();

                System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(dbConn.dbCntStr);//add connectionstring here

                // ContainerNo Linecode7   PackBarcode Quantity
                bcp.DestinationTableName = "ExtraContainerExtract";//give destination table name

                bcp.ColumnMappings.Add("ContainerNo", "ContainerNo");//Map all columns
                bcp.ColumnMappings.Add("Linecode7", "Linecode7");
                bcp.ColumnMappings.Add("PackBarcode", "PackBarcode");
                bcp.ColumnMappings.Add("Quantity", "Quantity");

                // and so on...., maap all source table with your destination table
                if (DtSource.Rows.Count > 0)
                {
                    bcp.WriteToServer(DtSource);
                }
            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
            }
        }

        #endregion ImportExtraContainerExtract


        #region DeletePOGrnHeader
        /// <summary>
        /// DeletePOGrnHeader
        /// </summary>
        /// <param name="dtSource"></param>
        public void DeletePOGrnHeader()
        {

            try
            {

                DatabaseConnectionDC dbConn = new DatabaseConnectionDC();

                SqlConnection cnTest = dbConn.CnDbConnection;
                SqlCommand cmdTest = new SqlCommand("delete [POGRNHeader] where [GRNNo]='" + GrnNo + "'", cnTest);
                cmdTest.ExecuteNonQuery();

            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
            }
        }

        #endregion DeletePOGrnHeader


        #region DeleteAllocationHeader
        /// <summary>
        /// DeleteAllocationHeader
        /// </summary>
        /// <param name="dtSource"></param>
        public void DeleteAllocationHeader()
        {

            try
            {

                DatabaseConnectionDC dbConn = new DatabaseConnectionDC();

                SqlConnection cnTest = dbConn.CnDbConnection;
                SqlCommand cmdTest = new SqlCommand("delete [AllocationHeader] where [AllocationNo]='" + AllocationNo + "'", cnTest);
                cmdTest.ExecuteNonQuery();

            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
            }
        }

        #endregion DeleteAllocationHeader


        #region InsertExtraContainerExtract
        /// <summary>
        /// InsertExtraContainerExtract
        /// </summary>
        /// <returns></returns>
        public bool InsertExtraContainerExtract()
        {
            bool Result = false;

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("InsertExtraContainerExtract", cnTest);

                cmdTest.Parameters.Add(new SqlParameter("@ContainerNo", SqlDbType.VarChar)).Value = ContainerNo;

                cmdTest.CommandType = CommandType.StoredProcedure;
                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertExtraContainerExtract


        #region ImportAdjustment
        /// <summary>
        /// ImportAdjustment
        /// </summary>
        /// <param name="dtSource"></param>
        public void ImportAdjustment()
        {

            try
            {

                DatabaseConnectionDC dbConn = new DatabaseConnectionDC();

                SqlConnection cnTest = dbConn.CnDbConnection;
                //SqlCommand cmdTest = new SqlCommand("Delete[ContainerExtract] where [Container Reference] = '"+ContainerNo+"'", cnTest);
                //cmdTest.ExecuteNonQuery();

                System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(dbConn.dbCntStr);//add connectionstring here

                //[AdjustmentNo] [AdjustmentDate] [LineNo] [LineCode7] [AdjustmentQty]
                bcp.DestinationTableName = "AdjustmentDetail";//give destination table name

                bcp.ColumnMappings.Add("AdjustmentNo", "AdjustmentNo");//Map all columns
                bcp.ColumnMappings.Add("OrgDocNo", "OrgDocNo");
                bcp.ColumnMappings.Add("AdjustmentDate", "AdjustmentDate");
                bcp.ColumnMappings.Add("LineNo", "LineNo");

                bcp.ColumnMappings.Add("PackBarcode", "PackBarcode");
                bcp.ColumnMappings.Add("AdjustmentQty", "AdjustmentQty");

                // and so on...., maap all source table with your destination table
                if (DtSource.Rows.Count > 0)
                {
                    bcp.WriteToServer(DtSource);
                }
            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
            }
        }
        #endregion ImportAdjustment


        #region InsertAdjustmentHeader
        /// <summary>
        /// InsertAdjustmentHeader
        /// </summary>
        /// <returns></returns>
        public bool InsertAdjustmentHeader()
        {
            bool Result = false;

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("InsertAdjustmentHeader", cnTest);

                cmdTest.Parameters.Add(new SqlParameter("@AdjNumber", SqlDbType.VarChar)).Value = AdjustmentNo;
                cmdTest.Parameters.Add(new SqlParameter("@DocNo", SqlDbType.VarChar)).Value = DocNo;
                cmdTest.Parameters.Add(new SqlParameter("@FileName", SqlDbType.VarChar)).Value = FileName;
                cmdTest.Parameters.Add(new SqlParameter("@CreatedId", SqlDbType.VarChar)).Value = "Admin";


                cmdTest.CommandType = CommandType.StoredProcedure;
                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
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
            DataTable dtTest = new DataTable();

            DatabaseConnectionDC dbTest = new DatabaseConnectionDC();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("CheckAdjustmentDetail", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@DocNo", SqlDbType.VarChar)).Value = AdjustmentNo;

                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetCheckAdjustmentDetails


        #endregion DCStock APP


        #region InsertTransferHeader
        /// <summary>
        /// InsertTransferHeader
        /// </summary>
        /// <returns></returns>
        public bool InsertTransferHeader()
        {
            bool Result = false;

            DatabaseConnection dbTest = new DatabaseConnection();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("InsertTransferHeader", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@DocNo", SqlDbType.VarChar)).Value = DocNo;
                cmdTest.Parameters.Add(new SqlParameter("@SenderFileName", SqlDbType.VarChar)).Value = FileName;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }
        #endregion InsertTransferHeader

        #region ImportTransferOrder
        /// <summary>
        /// TransferOrder
        /// </summary>
        /// <param name="dtSource"></param>
        public void ImportTransferOrder()
        {
            try
            {

                DatabaseConnection dbConn = new DatabaseConnection();

                SqlConnection cnTest = dbConn.CnDbConnection;
                //SqlCommand cmdTest = new SqlCommand("Delete[ContainerExtract] where [Container Reference] = '"+ContainerNo+"'", cnTest);
                //cmdTest.ExecuteNonQuery();

                System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(dbConn.dbCntStr);//add connectionstring here

                //[Barcode] [Quantity] 
                bcp.DestinationTableName = "TransferOrder";//give destination table name

                bcp.ColumnMappings.Add("DocNo", "DocNo");
                bcp.ColumnMappings.Add("Barcode", "Barcode");
                bcp.ColumnMappings.Add("Quantity", "TransferQty");
                bcp.ColumnMappings.Add("Location", "Location");

                bcp.ColumnMappings.Add("R", "ReceiverLocation");
                bcp.ColumnMappings.Add("Country", "Country");

                // and so on...., maap all source table with your destination table
                if (DtSource.Rows.Count > 0)
                {
                    bcp.WriteToServer(DtSource);
                }
            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
            }
        }
        #endregion ImportTransferOrder

        #region UpdateTransferOrder
        /// <summary>
        /// UpdateTransferOrder
        /// </summary>
        /// <returns></returns>
        public bool UpdateTransferOrder()
        {
            bool Result = false;

            DatabaseConnection dbTest = new DatabaseConnection();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("UpdateTransferOrder", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@DocNo", SqlDbType.VarChar)).Value = DocNo;
                cmdTest.Parameters.Add(new SqlParameter("@StoreNo", SqlDbType.VarChar)).Value = Location;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion UpdateTransferOrder


        #region UpdateReceivedQty
        /// <summary>
        /// UpdateReceivedQty
        /// </summary>
        /// <returns></returns>
        public bool UpdateReceivedQty()
        {
            bool Result = false;

            DatabaseConnection dbTest = new DatabaseConnection();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("UpdateReceivedQty", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@DocNo", SqlDbType.VarChar)).Value = DocNo;
                cmdTest.Parameters.Add(new SqlParameter("@Qty", SqlDbType.Decimal)).Value = Quantity;
                cmdTest.Parameters.Add(new SqlParameter("@Barcode", SqlDbType.VarChar)).Value = PackBarcode;
                cmdTest.Parameters.Add(new SqlParameter("@FileName", SqlDbType.VarChar)).Value = FileName;

                cmdTest.CommandTimeout = 0;

                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
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
            DataTable dtTest = new DataTable();

            DatabaseConnection dbTest = new DatabaseConnection();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("GetTransferOrder", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@DocNo", SqlDbType.VarChar)).Value = DocNo;
                cmdTest.CommandTimeout = 0;
                SqlDataAdapter daStock = new SqlDataAdapter(cmdTest);

                try
                {
                    daStock.Fill(dtTest);
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return dtTest;
        }
        #endregion GetTransferOrder

        #region UpdateTransferAdjustment
        /// <summary>
        /// UpdateTransferAdjustment
        /// </summary>
        /// <returns></returns>
        public bool UpdateTransferAdjustment()
        {
            bool Result = false;

            DatabaseConnection dbTest = new DatabaseConnection();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("UpdateTransferAdjustment", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@DocNo", SqlDbType.VarChar)).Value = DocNo;
                cmdTest.Parameters.Add(new SqlParameter("@Remarks", SqlDbType.VarChar)).Value = Remarks;
                cmdTest.Parameters.Add(new SqlParameter("@Qty", SqlDbType.Decimal)).Value = Quantity;
                cmdTest.Parameters.Add(new SqlParameter("@Id", SqlDbType.Int)).Value = Id;

                cmdTest.CommandTimeout = 0;
                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                    ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion UpdateTransferAdjustment

        #region InsertInventoryNAV
        /// <summary>
        /// InsertInventoryNAV
        /// </summary>
        /// <returns></returns>
        public bool InsertInventoryNAV()
        {
            bool Result = false;

            DatabaseConnection dbTest = new DatabaseConnection();

            if (dbTest.DbConnectionFailureMessage.Trim().Length <= 0)
            {
                SqlConnection cnTest = dbTest.CnDbConnection;

                SqlCommand cmdTest = new SqlCommand("InsertInventoryNAV", cnTest);
                cmdTest.CommandType = CommandType.StoredProcedure;

                cmdTest.Parameters.Add(new SqlParameter("@CompanyName", SqlDbType.VarChar)).Value = CompanyName;
                cmdTest.Parameters.Add(new SqlParameter("@LocationCode", SqlDbType.VarChar)).Value = Location;
                cmdTest.Parameters.Add(new SqlParameter("@DocNo", SqlDbType.VarChar)).Value = DocNo;
                

                cmdTest.CommandTimeout = 0;
                try
                {
                    Result = (cmdTest.ExecuteNonQuery() > 0) ? true : false;
                    ExceptionMessage = "";
                }
                catch (SqlException ex)
                {
                    ExceptionMessage = ex.Message;
                }

                finally
                {
                    dbTest.CloseDbConnection();
                }
            }
            else
            {
                ExceptionMessage = dbTest.DbConnectionFailureMessage.Trim();
            }

            return Result;
        }

        #endregion InsertInventoryNAV

        #region ImportHSCode
        /// <summary>
        /// ImportHSCode
        /// </summary>
        /// <param name="dtSource"></param>
        public void ImportHSCode()
        {

            try
            {

                DatabaseConnectionDC dbConn = new DatabaseConnectionDC();

                SqlConnection cnTest = dbConn.CnDbConnection;
                //SqlCommand cmdTest = new SqlCommand("Delete[ContainerExtract] where [Container Reference] = '"+ContainerNo+"'", cnTest);
                //cmdTest.ExecuteNonQuery();

                System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(dbConn.dbCntStr);//add connectionstring here

                //[AdjustmentNo] [AdjustmentDate] [LineNo] [LineCode7] [AdjustmentQty]
                bcp.DestinationTableName = "HSCodeMapping";//give destination table name

                bcp.ColumnMappings.Add("MatalanUKHSCode", "MatalanUKHSCode");//Map all columns
                bcp.ColumnMappings.Add("GccHSCode", "GccHSCode");
                bcp.ColumnMappings.Add("Description", "Description");
                //bcp.ColumnMappings.Add("IsActive", "IsActive");

                // and so on...., maap all source table with your destination table
                if (DtSource.Rows.Count > 0)
                {
                    bcp.WriteToServer(DtSource);
                }
            }
            catch (SqlException ex)
            {
                ExceptionMessage = ex.Message;
            }
        }
        #endregion ImportHSCode


        

    }
}