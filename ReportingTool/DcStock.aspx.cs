#region NameSpace
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Test.BAL;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using Excel;
using System.IO;

#endregion NameSpace

namespace ReportingTool
{
    public partial class DcStock : System.Web.UI.Page
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

        #region ddlType_SelectedIndexChanged
        /// <summary>
        /// ddlType_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddlType_SelectedIndexChanged(object sender, EventArgs e)
        {

            int CmdType = Convert.ToInt32(ddlType.SelectedItem.Value);

            GetStockDetails objDCStock = new GetStockDetails();

            DataTable dtData = null;

            switch (CmdType)
            {
                case 1:
                       BindPONumber();
                       trDocNo.Visible = true;
                       trStockLedger.Visible = false;
                       break;
                case 2:
                       BindSONumber();
                       trDocNo.Visible = true;
                       trStockLedger.Visible = false;
                       break;
                case 3:trStockLedger.Visible = true;
                       trDocNo.Visible = false;
                       break;
            }
            gdvSOData.Visible = false;
            gridView.Visible = false;

            btnDeleteSO.Visible = false;
            btnDeletePO.Visible = false;
        }
        #endregion ddlType_SelectedIndexChanged

        #region ddlPONumber_SelectedIndexChanged
        /// <summary>
        /// ddlPONumber_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddlPONumber_SelectedIndexChanged(object sender, EventArgs e)
        {

            int CmdType = Convert.ToInt32(ddlType.SelectedItem.Value);

            GetStockDetails objDCStock = new GetStockDetails();

            DataTable dtData = null;

            BindGrid();

          

        }
        #endregion ddlPONumber_SelectedIndexChanged


        #region btnLedger_Click
        /// <summary>
        /// btnLedger_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnLedger_Click(object sender, EventArgs e)
        {
            Boolean fileOK = false;
            Boolean fileFormat = false;
            String Msg = ""; ;
            String path = Server.MapPath("~/FileImport/");
            bool Result = false;
            if (IsPostBack)
            {

                if (fudStockLedger.HasFile)
                {
                    String fileExtension =
                        System.IO.Path.GetExtension(fudStockLedger.FileName).ToLower();
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
                        String fileName = "StockLedger" + DateTime.Now.Day + "-" + DateTime.Now.Month + "-" + DateTime.Now.Year + "_" + rnd.Next() + ".xlsx";
                        fudStockLedger.PostedFile.SaveAs(path
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

                        GetStockDetails objDC = new GetStockDetails();

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

                        // for Updating Stock Ledger
                        for (int i = 0; i < DtSource.Rows.Count; i++)
                        {

                            objDC.LineCode7 = DtSource.Rows[i]["LineCode7"].ToString();
                            objDC.PackBarcode = DtSource.Rows[i]["PackBarCode"].ToString();

                            objDC.PackID =DtSource.Rows[i]["PackID"].ToString();
                            objDC.PackType = DtSource.Rows[i]["PackType"].ToString();
                            objDC.LineCode7Qty = Convert.ToDecimal(DtSource.Rows[i]["LineCode7Qty"].ToString());
                            objDC.Outer = Convert.ToDecimal(DtSource.Rows[i]["Outer"].ToString());

                            objDC.PackLevel = DtSource.Rows[i]["PackLevel"].ToString();
                            objDC.ID = Convert.ToInt32(DtSource.Rows[i]["ID"].ToString());
                            Result = objDC.UpdateStockLedger();
                        }

                      
                        //6. Free resources (IExcelDataReader is IDisposable)
                        excelReader.Close();

                        if (Result)
                        {
                            lblMessage.ForeColor = System.Drawing.Color.Green;
                            lblMessage.Text = "Successfully Updated!";
                        }
                        else if (Msg.Length > 0)
                        {
                            lblMessage.ForeColor = System.Drawing.Color.Red;
                            lblMessage.Text = Msg;
                        }
                        else
                        {
                            lblMessage.ForeColor = System.Drawing.Color.Red;
                            lblMessage.Text = "Failed To Import Stock Ledger Data!";
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
        #endregion btnLedger_Click


        #region gridView_RowEditing
        /// <summary>
        /// gridView_RowEditing
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void gridView_RowEditing(object sender, GridViewEditEventArgs e)
        {
            gridView.EditIndex = e.NewEditIndex;
            BindGrid();
            //loadStores();
        }
        #endregion gridView_RowEditing

        #region gridView_RowUpdating
        /// <summary>
        /// gridView_RowUpdating
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void gridView_RowUpdating(object sender, GridViewUpdateEventArgs e)
        {
            
            //TextBox stor_name = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtname");
            //TextBox stor_address = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtaddress");
            //TextBox city = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtcity");
            //TextBox state = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtstate");
            //TextBox zip = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtzip");
            //con.Open();
            //SqlCommand cmd = new SqlCommand("update stores set stor_name='" + stor_name.Text + "', stor_address='" + stor_address.Text + "', city='" + city.Text + "', state='" + state.Text + "', zip='" + zip.Text + "' where stor_id=" + stor_id, con);
            //cmd.ExecuteNonQuery();
            //con.Close();
            //lblmsg.BackColor = Color.Blue;
            //lblmsg.ForeColor = Color.White;
            //lblmsg.Text = stor_id + "        Updated successfully........    ";
            //gridView.EditIndex = -1;
            //loadStores();

            UpdatePODetail(e);
            string ID = gridView.DataKeys[e.RowIndex].Values["ID"].ToString();
            lblMessage.BackColor = Color.Green;
            lblMessage.ForeColor = Color.White;
            lblMessage.Text = ID + "        Updated successfully........    ";
            gridView.EditIndex = -1;
            BindGrid();

        }
        #endregion gridView_RowUpdating

        #region gridView_RowCancelingEdit
        /// <summary>
        /// gridView_RowCancelingEdit
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void gridView_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
           gridView.EditIndex = -1;
            BindGrid();
        }
        #endregion gridView_RowCancelingEdit

        #region gridView_RowDeleting
        /// <summary>
        /// gridView_RowDeleting
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void gridView_RowDeleting(object sender, GridViewDeleteEventArgs e)
        {
            //string stor_id = gridView.DataKeys[e.RowIndex].Values["stor_id"].ToString();
            //con.Open();
            //SqlCommand cmd = new SqlCommand("delete from stores where stor_id=" + stor_id, con);
            //int result = cmd.ExecuteNonQuery();
            //con.Close();
            //if (result == 1)
            //{
            //    loadStores();
            //    lblmsg.BackColor = Color.Red;
            //    lblmsg.ForeColor = Color.White;
            //    lblmsg.Text = stor_id + "      Deleted successfully.......    ";
            //}
        }
        #endregion gridView_RowDeleting

        #region gridView_RowDataBound
        /// <summary>
        /// gridView_RowDataBound
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void gridView_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                string ID = Convert.ToString(DataBinder.Eval(e.Row.DataItem, "ID"));
                Button lnkbtnresult = (Button)e.Row.FindControl("ButtonDelete");
                if (lnkbtnresult != null)
                {
                    lnkbtnresult.Attributes.Add("onclick", "javascript:return deleteConfirm('" + ID + "')");
                }
            }
        }
        #endregion gridView_RowDataBound

        #region gridView_RowCommand
        /// <summary>
        /// gridView_RowCommand
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void gridView_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            if (e.CommandName.Equals("AddNew"))
            {
                //TextBox instorid = (TextBox)gridView.FooterRow.FindControl("instorid");
                //TextBox inname = (TextBox)gridView.FooterRow.FindControl("inname");
                //TextBox inaddress = (TextBox)gridView.FooterRow.FindControl("inaddress");
                //TextBox incity = (TextBox)gridView.FooterRow.FindControl("incity");
                //TextBox instate = (TextBox)gridView.FooterRow.FindControl("instate");
                //TextBox inzip = (TextBox)gridView.FooterRow.FindControl("inzip");
                //con.Open();
                //SqlCommand cmd =
                //    new SqlCommand(
                //        "insert into stores(stor_id,stor_name,stor_address,city,state,zip) values('" + instorid.Text + "','" +
                //        inname.Text + "','" + inaddress.Text + "','" + incity.Text + "','" + instate.Text + "','" + inzip.Text + "')", con);
                //int result = cmd.ExecuteNonQuery();
                //con.Close();
                //if (result == 1)
                //{
                //    loadStores();
                //    lblmsg.BackColor = Color.Green;
                //    lblmsg.ForeColor = Color.White;
                //    lblmsg.Text = instorid.Text + "      Added successfully......    ";
                //}
                //else
                //{
                //    lblmsg.BackColor = Color.Red;
                //    lblmsg.ForeColor = Color.White;
                //    lblmsg.Text = instorid.Text + " Error while adding row.....";
                //}
                InsertPODetail(e);
            }
        }
        #endregion gridView_RowCommand


        #region btnDeletePO_Click
        /// <summary>
        /// btnDeletePO_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDeletePO_Click(object sender, EventArgs e)
        {
            GetStockDetails ObjStock = new GetStockDetails();
            ObjStock.PONumber = ddlPONumber.SelectedItem.Value;
            bool Result=ObjStock.DeletePODetail();

            if(Result)
            {
                string ID = ddlPONumber.SelectedItem.Value;
                lblMessage.BackColor = Color.Green;
                lblMessage.ForeColor = Color.White;
                lblMessage.Text = ID + "        Deleted successfully........    ";
            }
        }
        #endregion btnDeletePO_Click


        #region btnDeleteSO_Click
        /// <summary>
        /// btnDeleteSO_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDeleteSO_Click(object sender, EventArgs e)
        {
            GetStockDetails ObjStock = new GetStockDetails();
            ObjStock.DocNo = ddlPONumber.SelectedItem.Value;
            bool Result = ObjStock.DeleteSODetail();

            if (Result)
            {
                string ID = ddlPONumber.SelectedItem.Value;
                lblMessage.BackColor = Color.Green;
                lblMessage.ForeColor = Color.White;
                lblMessage.Text = ID + "        Deleted successfully........    ";
            }
        }
        #endregion btnDeleteSO_Click

        #endregion Events

        #region Methods

        #region BindPONumber
        /// <summary>
        /// BindPONumber
        /// </summary>
        private void BindPONumber()
        {
            GetStockDetails objDCStock = new GetStockDetails();

            DataTable dtData = objDCStock.GetPOHeader();

            ddlPONumber.DataSource = dtData;
            ddlPONumber.DataMember = "PONumber";
            ddlPONumber.DataValueField = "PONumber";
            ddlPONumber.DataBind();

           // ddlPONumber.Text = "Select";
            ddlPONumber.SelectedItem.Value = "Select";

        }
        #endregion BindPONumber

        #region BindSONumber
        /// <summary>
        /// BindSONumber
        /// </summary>
        private void BindSONumber()
        {
            GetStockDetails objDCStock = new GetStockDetails();

            DataTable dtData = objDCStock.GetSOHeader();

            ddlPONumber.DataSource = dtData;
            ddlPONumber.DataMember = "SONumber";
            ddlPONumber.DataValueField = "SONumber";
            ddlPONumber.DataBind();
            // ddlPONumber.Text = "Select";
            ddlPONumber.SelectedItem.Value = "Select";

        }
        #endregion BindSONumber
        
        #region BindGrid
        /// <summary>
        /// Bind Grid
        /// </summary>
        private void BindGrid()
        {
            int CmdType = Convert.ToInt32(ddlType.SelectedItem.Value);

            GetStockDetails objDCStock = new GetStockDetails();

            DataTable dtData = null;

            switch (CmdType)
            {
                case 1:
                    objDCStock.PONumber = ddlPONumber.SelectedItem.Text;
                    dtData = objDCStock.GetPODetail();
                    if (dtData.Rows.Count > 0)
                    {
                        if (IsPostBack)
                        {
                            gridView.DataSource = dtData;
                            gridView.DataBind();
                            gdvSOData.Visible = false;
                            gridView.Visible = true;
                        }
                        btnDeletePO.Visible = true;
                        btnDeleteSO.Visible = false;
                        lblMessage.Text = "";
                    }
                    else
                    {
                        lblMessage.Text = "No Data Found!";
                        lblMessage.ForeColor = Color.Red;
                        gdvSOData.Visible = false;
                        gridView.Visible = false;
                        btnDeletePO.Visible = false;
                        btnDeleteSO.Visible = false;
                    }
                    break;
                case 2:
                    objDCStock.DocNo = ddlPONumber.SelectedItem.Text;
                    dtData = objDCStock.GetSODetail();
                    if (dtData.Rows.Count > 0)
                    {
                        if (IsPostBack)
                        {
                            gdvSOData.DataSource = dtData;
                            gdvSOData.DataBind();
                            gdvSOData.Visible = true;
                            gridView.Visible = false;

                            btnDeleteSO.Visible = true;
                            btnDeletePO.Visible = false;
                        }

                        lblMessage.Text = "";
                    }
                    else
                    {
                        lblMessage.Text = "No Data Found!";
                        lblMessage.ForeColor = Color.Red;
                        gdvSOData.Visible = false;
                        gridView.Visible = false;

                        btnDeletePO.Visible = false;
                        btnDeleteSO.Visible = false;
                    }
                    break;
            }

        }
        #endregion BindGrid
        
        #region UpdatePODetail
        /// <summary>
        /// Update PODetail
        /// </summary>
        private void UpdatePODetail(GridViewUpdateEventArgs e)
        {
            
            int ID = Convert.ToInt32(gridView.DataKeys[e.RowIndex].Values["ID"].ToString());
            TextBox LineCode7 = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtLineCode7");
            TextBox PackID = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtPackID");
            TextBox PackBarcode = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtPackBarcode");

            TextBox PackType = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtPackType");
            TextBox OrderQty = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtOrderQty");
            TextBox UnitPrice = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtUnitPrice");
            TextBox COO = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtCOO");

            TextBox Department = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtDepartment");
            TextBox Nest = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtNest");
            TextBox Description = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtDescription");
            TextBox Season = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtSeason");

            TextBox Outer = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtOuter");
            TextBox Invoiced = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtInvoiced");
            TextBox PackLevel = (TextBox)gridView.Rows[e.RowIndex].FindControl("txtPackLevel");


            GetStockDetails ObjDC = new GetStockDetails();

            ObjDC.LineCode7 = LineCode7.Text;
            ObjDC.PackID = PackID.Text;
            ObjDC.PackBarcode = PackBarcode.Text;
            ObjDC.PackType = PackType.Text;

            ObjDC.OrderQty = Convert.ToDecimal(OrderQty.Text);
            ObjDC.UnitPrice = Convert.ToDecimal(UnitPrice.Text);
            ObjDC.COO = COO.Text;
            ObjDC.Department = Department.Text;

            ObjDC.Nest = Nest.Text;
            ObjDC.Description = Description.Text;
            ObjDC.Season = Season.Text;
            ObjDC.Outer = Convert.ToDecimal(Outer.Text);

            ObjDC.Invoiced = Convert.ToDecimal(Invoiced.Text);
            ObjDC.PackLevel = PackLevel.Text;
            ObjDC.ID = ID;

            bool Result =ObjDC.UpdatePODetail();

        }
        #endregion UpdatePODetail

        #region InsertPODetail
        /// <summary>
        /// Insert PODetail
        /// </summary>
        private void InsertPODetail(GridViewCommandEventArgs e)
        {
                       
            TextBox LineCode7 = (TextBox)gridView.FooterRow.FindControl("inLineCode7");
            TextBox PackID = (TextBox)gridView.FooterRow.FindControl("inPackID");
            TextBox PackBarcode = (TextBox)gridView.FooterRow.FindControl("inPackBarcode");

            TextBox PackType = (TextBox)gridView.FooterRow.FindControl("inPackType");
            TextBox OrderQty = (TextBox)gridView.FooterRow.FindControl("inOrderQty");
            TextBox UnitPrice = (TextBox)gridView.FooterRow.FindControl("inUnitPrice");
            TextBox COO = (TextBox)gridView.FooterRow.FindControl("inCOO");

            TextBox Department = (TextBox)gridView.FooterRow.FindControl("inDepartment");
            TextBox Nest = (TextBox)gridView.FooterRow.FindControl("inNest");
            TextBox Description = (TextBox)gridView.FooterRow.FindControl("inDescription");
            TextBox Season = (TextBox)gridView.FooterRow.FindControl("inSeason");

            TextBox Outer = (TextBox)gridView.FooterRow.FindControl("inOuter");
            TextBox Invoiced = (TextBox)gridView.FooterRow.FindControl("inInvoiced");
            TextBox PackLevel = (TextBox)gridView.FooterRow.FindControl("inPackLevel");


            GetStockDetails ObjDC = new GetStockDetails();

            ObjDC.PONumber = ddlPONumber.Text.Trim();
            ObjDC.LineCode7 = LineCode7.Text;
            ObjDC.PackID = PackID.Text;
            ObjDC.PackBarcode = PackBarcode.Text;

            ObjDC.PackType = PackType.Text;
            ObjDC.OrderQty = Convert.ToDecimal(OrderQty.Text);
            ObjDC.UnitPrice = Convert.ToDecimal(UnitPrice.Text);
            ObjDC.COO = COO.Text;

            ObjDC.Department = Department.Text;
            ObjDC.Nest = Nest.Text;
            ObjDC.Description = Description.Text;
            ObjDC.Season = Season.Text;

            ObjDC.Outer = Convert.ToDecimal(Outer.Text);
            ObjDC.Invoiced = Convert.ToDecimal(Invoiced.Text);
            ObjDC.PackLevel = PackLevel.Text;
           // ObjDC.ID = ID;

            bool Result = ObjDC.InsertPODetail();

            if (Result == true)
            {
                BindGrid();
                lblMessage.BackColor = Color.Green;
                lblMessage.ForeColor = Color.White;
                lblMessage.Text =  " New Line Added successfully......    ";
            }
            else
            {
                lblMessage.BackColor = Color.Red;
                lblMessage.ForeColor = Color.White;
                lblMessage.Text = " Error while adding row.....";
            }
        }


        #endregion InsertPODetail

        #endregion Methods

        
    }
}