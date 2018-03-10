
#region NameSpace
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Test.BAL;
#endregion NameSpace

namespace Test
{
    public partial class Update : System.Web.UI.Page
    {
        public const int UpdateTableProcessId = 1; 
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

        #region btnUpdate_Click
        /// <summary>
        /// btnUpdate_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnUpdate_Click(object sender, EventArgs e)
        {

            if (GetProcessStatus())
            {
                lblMessage.ForeColor = System.Drawing.Color.Red;
                lblMessage.Text = "Tables Are Locked By Another User,Try Again Later";
            }
            else
            {
                if (Rdltem.SelectedValue == "0")
                {
                    UpdateOfferPrice();
                }
                UpdateTables();
            }

            Page.ClientScript.RegisterStartupScript(this.GetType(), "CallMyFunction", "$('#btnUpdate').Show();", true);
        }
        #endregion btnUpdate_Click

        #region Rdltem_SelectedIndexChanged
        /// <summary>
        ///  Rdltem_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void Rdltem_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (Rdltem.SelectedValue == "1")
                trPromotion.Visible = false;
            else
                trPromotion.Visible = true;
        }
        #endregion Rdltem_SelectedIndexChanged
       

        #endregion Events

        #region Methods

        #region UpdateTables
        /// <summary>
        /// Update Tables
        /// </summary>
        private void UpdateTables()
        {
            GetStockDetails ObjStock = new GetStockDetails();

            ObjStock.ItemOperationType = (Rdltem.SelectedValue == "1") ? true : false;
            //ObjStock.ILEOperationType = (RdlItemLedger.SelectedValue == "1") ? true : false;
            
            if (RdlItemLedger.SelectedValue == "1")
                ObjStock.ILEOperationType = 1;
            else if(RdlItemLedger.SelectedValue == "0")
                ObjStock.ILEOperationType = 0;
            else
                ObjStock.ILEOperationType = 2;

            if (RdlValueEntry.SelectedValue == "1")
                ObjStock.ValueOperationType = 1;
            else if (RdlValueEntry.SelectedValue == "0")
                ObjStock.ValueOperationType = 0;
            else
                ObjStock.ValueOperationType = 2;

            
            
            ObjStock.FootFallOperationType = (RdlFootFall.SelectedValue == "1") ? true : false;
            ObjStock.TransactionOperationType = (RdlTransHeader.SelectedValue == "1") ? true : false;
            // ObjStock.ValueOperationType = (RdlValueEntry.SelectedValue == "1") ? true : false;
                      

            bool Result = ObjStock.UpdateTables();

           
            if (Result == true)
            {
                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Successfuly Completed.";
            }
            else
            {
                lblMessage.Visible = true;
                lblMessage.ForeColor = System.Drawing.Color.Red;
                lblMessage.Text = "Tables Updation Failed.";
            }
        }
        #endregion UpdateTables

        #region UpdateOfferPrice
        /// <summary>
        /// Update Offer Price
        /// </summary>
        private void UpdateOfferPrice()
        {
            GetStockDetails objStock = new GetStockDetails();
            
            objStock.BahrainOffer = txtBahrainOffer.Text.Trim();
            objStock.UaeOffer = txtUaeOffer.Text.Trim();
            objStock.OmanOffer = txtOmanOffer.Text.Trim();
            objStock.JordanOffer = txtJordanOffer.Text.Trim();
            objStock.QatarOffer = txtQatarOffer.Text.Trim();
            objStock.KsaOffer = txtKsaOffer.Text.Trim();
            bool Result=objStock.UpdateOfferPrice();

        }
        #endregion UpdateOfferPrice


        #region GetProcessStatus
        /// <summary>
        /// Get Process Status
        /// </summary>
        private bool GetProcessStatus()
        {
            GetStockDetails objStock = new GetStockDetails();
            objStock.ProcessStatusId = UpdateTableProcessId;
            DataTable dtStatus=objStock.GetProcessStatus();
            bool Flag = Convert.ToBoolean(dtStatus.Rows[0]["Flag"]);

            return Flag;
        }
        #endregion GetProcessStatus

        #endregion Methods

    }
}