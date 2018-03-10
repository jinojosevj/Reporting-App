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
    public partial class MatalanSalesAndStock : System.Web.UI.Page
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
            if (ddlType.SelectedItem.Value == "1")
                GenerateSalesData();
            else
                GenerateStockData();
        }
        #endregion btnGenerate_Click

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
                    sw.Write("|");
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
                        sw.Write("|");
                }
                sw.Write(sw.NewLine);
            }
            sw.Write(sw.NewLine);
        }

        #endregion Export To Csv

        #region GenerateSalesData
        /// <summary>
        /// GenerateSalesData
        /// </summary>
        private void GenerateSalesData()
        {
            GetStockDetails ObjStock = new GetStockDetails();
            ObjStock.AsOfDate = Convert.ToDateTime(txtAsOfDate.Text.Trim().ToString());
            //DateTime.ParseExact(txtAsOfDate.Text.Trim().ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            int SequenceNo = Convert.ToInt32(txtSequenceNo.Text.Trim());

            if (txtStoreNo.Text.Length > 0)
            {
                ObjStock.Location = txtStoreNo.Text.Trim();
                DataTable dt = ObjStock.GetSalesData();
                string filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                StreamWriter sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();
            }
            else
            {

                ObjStock.Location = "0400";
                DataTable dt = ObjStock.GetSalesData();
                string filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                StreamWriter sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0401";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0402";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0403";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0404";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();


                SequenceNo++;
                ObjStock.Location = "0405";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0406";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0407";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0409";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();


                SequenceNo++;
                ObjStock.Location = "0410";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();


                SequenceNo++;
                ObjStock.Location = "0411";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0412";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();


                SequenceNo++;
                ObjStock.Location = "0414";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0415";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0416";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0417";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0418";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0419";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0421";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0424";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0425";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0426";
                dt = ObjStock.GetSalesData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Sales\\" + "MESALES0" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();
            }
            lblMessage.Text = "Sales Files Generated";
            lblMessage.ForeColor = Color.Green;
        }
        #endregion GenerateSalesData

        #region GenerateStockData
        /// <summary>
        /// GenerateStockData
        /// </summary>
        private void GenerateStockData()
        {
            GetStockDetails ObjStock = new GetStockDetails();
            int SequenceNo = Convert.ToInt32(txtSequenceNo.Text.Trim());

            if (txtStoreNo.Text.Length > 0)
            {
                ObjStock.Location = txtStoreNo.Text.ToString();
                DataTable dt = ObjStock.GetStockData();
                string filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                StreamWriter sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();
            }
            else
            {
                ObjStock.Location = "0400";
                DataTable dt = ObjStock.GetStockData();
                string filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                StreamWriter sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0401";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0402";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0403";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0404";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();


                SequenceNo++;
                ObjStock.Location = "0405";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0406";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0407";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0409";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();


                SequenceNo++;
                ObjStock.Location = "0410";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();


                SequenceNo++;
                ObjStock.Location = "0411";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0412";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();


                SequenceNo++;
                ObjStock.Location = "0414";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0415";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0416";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0417";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0418";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0419";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0421";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0424";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0425";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();

                SequenceNo++;
                ObjStock.Location = "0426";
                dt = ObjStock.GetStockData();
                filePath = Server.MapPath(".") + "\\SalesAndStock\\Stock\\" + "MESYN00" + SequenceNo;
                sw = new StreamWriter(@filePath, false);
                ExportToCsv(dt, sw);
                sw.Close();
            }
            lblMessage.Text = "Stock Files Generated";
            lblMessage.ForeColor = Color.Green;
        }
        #endregion GenerateStockData

        #endregion Methods
    }
}