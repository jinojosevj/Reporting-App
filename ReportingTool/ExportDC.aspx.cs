#region NameSpace
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ReportingTool.BAL;
#endregion NameSpace

namespace ReportingTool
{
    public partial class ExportDC : System.Web.UI.Page
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
            BindPONumber();
            //BindStore();
            btnDownload.Visible = false;
            
        }
        #endregion Page_Load

        #region btnExport_Click
        /// <summary>
        /// btnExport_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnExport_Click(object sender, EventArgs e)
        {
            DataTable dt = null;
            TatiBAL objDC = new TatiBAL();

            objDC.PONumber = CmbPONumber.Text.Trim();
            objDC.ReportType = ddlType.SelectedItem.Value;

            if(ddlType.SelectedItem.Value=="12" || ddlType.SelectedItem.Value == "13")
            {
                objDC.SONumber = ddlSONumber.Text.Trim();
            }
            else
            {
                objDC.SONumber = txtSONumber.Text.Trim();
            }
            
            objDC.AllocationNo = txtAllocationNo.Text.Trim();

            objDC.Location = CmbStoreNo.Text.Trim();
            objDC.LineCode = txtLineCode7.Text.Trim();
            objDC.PackBarcode = txtPackBarcode.Text.Trim();
            objDC.IssueNo = txtIssueNo.Text.Trim();
            objDC.DocNo = "";

            string filePath = "";
            Random rnd = new Random();
            switch (ddlType.SelectedItem.Value)
            {
                case "1":
                          filePath = Server.MapPath(".") + "\\Reports\\MPO" + CmbPONumber.Text.Trim().Substring(CmbPONumber.Text.Length-5, 5) + "_" + rnd.Next() + ".csv";
                          break;
                case "2":
                          filePath = Server.MapPath(".") + "\\Reports\\ITEM" + CmbPONumber.Text.Trim().Substring(CmbPONumber.Text.Length - 5, 5) + "_" + rnd.Next() + ".csv";
                          break;
                case "3":
                          filePath = Server.MapPath(".") + "\\Reports\\ITEM_SNG" + CmbPONumber.Text.Trim().Substring(CmbPONumber.Text.Length - 5, 5) + "_" + rnd.Next() + ".csv";
                          break;
                case "4":
                          filePath = Server.MapPath(".") + "\\Reports\\MSO-" + txtSONumber.Text.Trim().Substring(txtSONumber.Text.Length - 5, 5) + "_" + rnd.Next() + ".csv";
                          break;
                case "5":
                          filePath = Server.MapPath(".") + "\\Reports\\ProductLineListing"  + "_" + rnd.Next() + ".csv";
                          break;
                case "6":
                         filePath = Server.MapPath(".") + "\\Reports\\FamilyListing" + "_" + rnd.Next() + ".csv";
                         break;
                case "7":
                         filePath = Server.MapPath(".") + "\\Reports\\StockLedgerByLinecode7" + "_" + rnd.Next() + ".csv";
                         break;
                case "8":
                         filePath = Server.MapPath(".") + "\\Reports\\StockLedgerByPackBarcode" + "_" + rnd.Next() + ".csv";
                         break;
                case "9":
                         filePath = Server.MapPath(".") + "\\Reports\\PackExtractByLinecode7" + "_" + rnd.Next() + ".csv";
                         break;
                case "10":
                         filePath = Server.MapPath(".") + "\\Reports\\PackExtractByPackBarcode" + "_" + rnd.Next() + ".csv";
                         break;
                case "11":
                        filePath = Server.MapPath(".") + "\\Reports\\IssueNoteLinecode12" + "_" + rnd.Next() + ".csv";
                        break;
                case "12":
                        filePath = Server.MapPath(".") + "\\Reports\\MSO-" + ddlSONumber.Text.Trim().Substring(ddlSONumber.Text.Trim().Length - 5, 5) + "_" + rnd.Next() + ".csv";
                        break;
                case "13":
                       filePath = Server.MapPath(".") + "\\Reports\\SO_Header" + "_" + rnd.Next() + ".csv";
                       break;

                case "14":
                         
                         if(rblInwardSelectFile.SelectedItem.Value=="1")
                           {
                             objDC.ReportType = "14";
                             objDC.DocNo = ddlContainerNo.Text;
                             filePath = Server.MapPath(".") + "\\Reports\\Container" + "_" + rnd.Next() + ".csv";
                           }

                           if (rblInwardSelectFile.SelectedItem.Value == "2")
                           {
                             objDC.ReportType = "15";
                             objDC.DocNo = ddlPONumber.Text;
                             filePath = Server.MapPath(".") + "\\Reports\\PO" + "_" + rnd.Next() + ".csv";
                           }

                            if (rblInwardSelectFile.SelectedItem.Value == "3")
                            {
                                objDC.ReportType = "16";
                                objDC.DocNo = ddlPOGrn.Text;
                                filePath = Server.MapPath(".") + "\\Reports\\POGRN" + "_" + rnd.Next() + ".csv";
                            }
                      break;

                case "17":
                    
                    if (rdlSelectDoc.SelectedItem.Value == "1" || rdlSelectDoc.SelectedItem.Value == "2" || rdlSelectDoc.SelectedItem.Value == "3")
                    {

                        if (rblInwardSelectFile.SelectedItem.Value == "1")
                        {
                            objDC.ReportType = "14";
                            objDC.DocNo = txtIWContainer.Text.Trim();
                            filePath = Server.MapPath(".") + "\\Reports\\Container" + "_" + rnd.Next() + ".csv";
                        }

                        if (rblInwardSelectFile.SelectedItem.Value == "2")
                        {
                            objDC.ReportType = "15";
                            objDC.DocNo = txtIWPO.Text.Trim();
                            filePath = Server.MapPath(".") + "\\Reports\\PO" + "_" + rnd.Next() + ".csv";
                        }

                        if (rblInwardSelectFile.SelectedItem.Value == "3")
                        {
                            objDC.ReportType = "16";
                            objDC.DocNo = txtIWGRN.Text.Trim();
                            filePath = Server.MapPath(".") + "\\Reports\\POGRN" + "_" + rnd.Next() + ".csv";
                        }
                    }
                    else
                    {
                        if (RdlOWSelectFile.SelectedItem.Value == "1")
                        {
                            objDC.ReportType = "17";
                            objDC.DocNo = txtOWAllocation.Text.Trim();
                            filePath = Server.MapPath(".") + "\\Reports\\Allocation" + "_" + rnd.Next() + ".csv";
                        }

                        if (RdlOWSelectFile.SelectedItem.Value == "2")
                        {
                            objDC.ReportType = "18";
                            objDC.DocNo = txtOWSO.Text.Trim();
                            filePath = Server.MapPath(".") + "\\Reports\\SO" + "_" + rnd.Next() + ".csv";
                        }

                        if (RdlOWSelectFile.SelectedItem.Value == "3")
                        {
                            objDC.ReportType = "19";
                            objDC.DocNo = txtOWIssueNote.Text.Trim();
                            filePath = Server.MapPath(".") + "\\Reports\\IssueNote" + "_" + rnd.Next() + ".csv";
                        }

                    }

                    break;

                case "20":

                    filePath = Server.MapPath(".") + "\\Reports\\PackExtract" + "_" + rnd.Next() + ".csv";
                    break;
                case "21":
                    filePath = Server.MapPath(".") + "\\Reports\\AllocationByPackbarcode" + "_" + rnd.Next() + ".csv";
                    break;
                case "22":
                    filePath = Server.MapPath(".") + "\\Reports\\AllocationByLineCode" + "_" + rnd.Next() + ".csv";
                    break;

                case "23":
                    filePath = Server.MapPath(".") + "\\Reports\\SOByPackbarcode" + "_" + rnd.Next() + ".csv";
                    break;
                case "24":
                    filePath = Server.MapPath(".") + "\\Reports\\SOByLineCode" + "_" + rnd.Next() + ".csv";
                    break;
                case "25":
                    objDC.DocNo = txtDocumentNo.Text.Trim();
                    filePath = Server.MapPath(".") + "\\Reports\\Adjustment" + "_" + rnd.Next() + ".csv";
                    break;
                case "26":
                    filePath = Server.MapPath(".") + "\\Reports\\HSCode" + "_" + rnd.Next() + ".csv";
                    break;
            }

            dt = objDC.GetPODetails();
            if (dt != null && dt.Rows.Count > 0)
            {
                ViewState["FileName"] = filePath;
                StreamWriter sw = new StreamWriter(@filePath, false);

                ExportToCsv(dt, sw);
                sw.Close();

                btnDownload.Visible = true;
                lblMessage.Text = "Report Generated";
                lblMessage.ForeColor = Color.Green;
            }
            else
            {
                ViewState["FileName"] = "";
                lblMessage.Text = "No Data Found!";
                lblMessage.ForeColor = Color.Red;
                btnDownload.Visible = false;
            }
        }
        #endregion btnExport_Click


        #region btnDownload_Click
        /// <summary>
        /// btnDownload_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnDownload_Click(object sender, EventArgs e)
        {
            string fileName = ViewState["FileName"].ToString();
            FileDownload(fileName);
        }
        #endregion btnDownload_Click


        #region ddlType_SelectedIndexChanged
        /// <summary>
        /// ddlType_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddlType_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblMessage.Text = "";
            btnDownload.Visible = false;

            switch(ddlType.SelectedItem.Value)
            {
                case "1":
                    trPONumber.Visible = true;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = false;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;
                    break;
                case "2":
                    trPONumber.Visible = true;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = false;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;
                    break;
                case "3":
                    trPONumber.Visible = true;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = false;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;
                    break;
                case "4":
                    trPONumber.Visible = false;
                    trSONumber.Visible = true;
                    trStoreNo.Visible = true;
                    trAllocationNo.Visible = true;
                    BindStore();
                    ddlAllocation.Visible = true;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = false;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;
                    break;
                case "5":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = true;
                    trAllocationNo.Visible = false;
                    BindStore();
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = false;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;
                    break;
                case "6":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = true;
                    trAllocationNo.Visible = false;
                    BindStore();
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = false;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;
                    break;
                case "7":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = true;
                    trIssueNo.Visible = false;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;
                    break;

                case "8":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = true;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = false;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false; 
                    break;
                case "9":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = true;
                    trIssueNo.Visible = false;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;
                    break;
                case "10":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = true;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = false;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;
                    break;
                case "11":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = true;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;

                    break;
                case "12":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = false;
                    BindSONumber();
                    trCmbSONumber.Visible = true;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;
                    break;
                case "13":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = false;
                    BindSONumber();
                    trCmbSONumber.Visible = true;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;
                    break;

                case "14":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = false;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = true;
                    trContainerNo.Visible = true;
                    trPOGrn.Visible = true;
                    trInwardSelectFile.Visible = true;

                    BindAllPONumber();
                    BindAllPOGrn();
                    BindAllContainer();

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible =false;
                    trDocumentNo.Visible = false;

                    break;
                case "17":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = false;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = true;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = true;
                    trSelectDoc.Visible = true;
                    trRdlOWSelectFile.Visible = false;
                    
                    trInwardDocs.Visible = true;
                    
                    BindDoc("1", "ContainerNo");
                    trDocumentNo.Visible = false;
                    break;

                case "20":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;

                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = false;

                    trCmbSONumber.Visible = false;
                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;

                    trInwardSelectFile.Visible = false;
                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;

                    trRdlOWSelectFile.Visible = false;
                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;
                    break;

                case "21":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = true;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = false;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;
                    break;
                case "22":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = true;
                    trIssueNo.Visible = false;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;
                    break;

                case "23":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = true;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = false;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;

                    break;
                case "24":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = true;
                    trIssueNo.Visible = false;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;
                    break;
                case "25":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;
                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = false;
                    trCmbSONumber.Visible = false;

                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;
                    trInwardSelectFile.Visible = false;

                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;
                    trRdlOWSelectFile.Visible = false;

                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = true;
                    break;

                case "26":
                    trPONumber.Visible = false;
                    trSONumber.Visible = false;
                    trStoreNo.Visible = false;
                    trAllocationNo.Visible = false;

                    ddlAllocation.Visible = false;
                    trPackBarcode.Visible = false;
                    trLineCode7.Visible = false;
                    trIssueNo.Visible = false;

                    trCmbSONumber.Visible = false;
                    trddlPONumber.Visible = false;
                    trContainerNo.Visible = false;
                    trPOGrn.Visible = false;

                    trInwardSelectFile.Visible = false;
                    trOutwardDocs.Visible = false;
                    trSearch.Visible = false;
                    trSelectDoc.Visible = false;

                    trRdlOWSelectFile.Visible = false;
                    trInwardDocs.Visible = false;
                    trDocumentNo.Visible = false;
                    break;

            }
        }
        #endregion ddlType_SelectedIndexChanged

        #region ddl_TextChanged
        /// <summary>
        /// ddl_TextChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddl_TextChanged(object sender, EventArgs e)
        {
            txtAllocationNo.Text = txtAllocationNo.Text + "," + ddlAllocation.SelectedItem.Text;
        }
        #endregion ddl_TextChanged


        #region CmbStoreNo_SelectedIndexChanged
        /// <summary>
        /// CmbStoreNo_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void CmbStoreNo_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtAllocationNo.Text = "";
            BindAllocation();
        }
        #endregion CmbStoreNo_SelectedIndexChanged


        #region ddlContainerNo_SelectedIndexChanged
        /// <summary>
        /// ddlContainerNo_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddlContainerNo_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblMessage.Text = "";
            btnDownload.Visible = false;

            TatiBAL objDC = new TatiBAL();
            objDC.ReportType = "3";
            objDC.DocNo = ddlContainerNo.Text;
            DataTable dt = objDC.GetAllPOHeader();

            if(dt.Rows.Count>0)
            {
                ddlPONumber.Text = dt.Rows[0]["PONumber"].ToString();
                ddlPOGrn.Text= dt.Rows[0]["GRNNumber"].ToString();
            }
            else
            {
                lblMessage.Text = "No Data Found!";
                lblMessage.ForeColor = Color.Red;
            }
        }
        #endregion ddlContainerNo_SelectedIndexChanged


        #region ddlPONumber_SelectedIndexChanged
        /// <summary>
        /// ddlPONumber_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ddlPONumber_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblMessage.Text = "";
            btnDownload.Visible = false;

            TatiBAL objDC = new TatiBAL();
            objDC.ReportType = "2";
            objDC.DocNo = ddlPONumber.Text;
            DataTable dt = objDC.GetAllPOHeader();

            if (dt.Rows.Count > 0)
            {
                ddlContainerNo.Text = dt.Rows[0]["ContainerNo"].ToString();
                ddlPOGrn.Text = dt.Rows[0]["GRNNumber"].ToString();
            }
            else
            {
                lblMessage.Text = "No Data Found!";
                lblMessage.ForeColor = Color.Red;
            }
        }
        #endregion ddlPONumber_SelectedIndexChanged



        #region ddlPOGrn_SelectedIndexChanged
        /// <summary>
        /// ddlPOGrn_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        protected void ddlPOGrn_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblMessage.Text = "";
            btnDownload.Visible = false;

            TatiBAL objDC = new TatiBAL();
            objDC.ReportType = "4";
            objDC.DocNo = ddlPOGrn.Text;
            DataTable dt = objDC.GetAllPOHeader();

            if (dt.Rows.Count > 0)
            {
                ddlContainerNo.Text = dt.Rows[0]["ContainerNo"].ToString();
                ddlPONumber.Text = dt.Rows[0]["PONumber"].ToString();
            }
            else
            {
                lblMessage.Text = "No Data Found!";
                lblMessage.ForeColor = Color.Red;
            }

        }
        #endregion ddlPOGrn_SelectedIndexChanged


        #region rdlSelectDoc_SelectedIndexChanged
        /// <summary>
        /// rdlSelectDoc_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void rdlSelectDoc_SelectedIndexChanged(object sender, EventArgs e)
        {

            txtIWContainer.Text = "";
            txtIWPO.Text = "";
            txtIWGRN.Text = "";

            txtOWAllocation.Text = "";
            txtOWIssueNote.Text = "";
            txtOWSO.Text = "";
            btnDownload.Visible = false;
            lblMessage.Text = "";

            switch (rdlSelectDoc.SelectedItem.Value)
            {
                case "1":
                        BindDoc("1", "ContainerNo");
                        trInwardDocs.Visible = true;
                        trOutwardDocs.Visible = false;
                        trInwardSelectFile.Visible = true;
                        trRdlOWSelectFile.Visible = false;
                        break;
                case "2":
                        BindDoc("1", "PONumber");
                        trInwardDocs.Visible = true;
                        trOutwardDocs.Visible = false;
                        trInwardSelectFile.Visible = true;
                        trRdlOWSelectFile.Visible = false;
                    break;
                case "3":
                        BindDoc("1", "GRNNumber");
                        trInwardDocs.Visible = true;
                        trOutwardDocs.Visible = false;
                        trInwardSelectFile.Visible = true;
                        trRdlOWSelectFile.Visible = false;
                    break;


                case "4":
                        BindDoc("5", "AllocationNo");
                        trInwardDocs.Visible = false;
                        trOutwardDocs.Visible = true;
                        trInwardSelectFile.Visible = false;
                        trRdlOWSelectFile.Visible = true;
                    break;
                case "5":
                        BindDoc("5", "SONumber");
                        trInwardDocs.Visible = false;
                        trOutwardDocs.Visible = true;
                        trInwardSelectFile.Visible = false;
                        trRdlOWSelectFile.Visible = true;
                    break;
                case "6":
                        BindDoc("5", "IssueNoteNo");
                        trInwardDocs.Visible = false;
                        trOutwardDocs.Visible = true;
                        trInwardSelectFile.Visible = false;
                        trRdlOWSelectFile.Visible = true;
                    break;
            }

        }
        #endregion rdlSelectDoc_SelectedIndexChanged


        #region cmbSearch_SelectedIndexChanged
        /// <summary>
        /// cmbSearch_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void cmbSearch_SelectedIndexChanged(object sender, EventArgs e)
        {
            TatiBAL objDC = new TatiBAL();
            DataTable dt = new DataTable();
            objDC.DocNo = cmbSearch.Text.Trim();

            switch (rdlSelectDoc.SelectedItem.Value)
            {
                case "1":
                    objDC.ReportType = "3";
                    break;
                case "2":
                    objDC.ReportType = "2";
                    break;
                case "3":
                    objDC.ReportType = "4";
                    break;
                case "4":
                    objDC.ReportType = "7";
                    break;
                case "5":
                    objDC.ReportType = "6";
                    break;
                case "6":
                    objDC.ReportType = "8";
                    break;
            }
                 dt = objDC.GetAllPOHeader();
                 if (dt.Rows.Count > 0)
                    {
                            if (rdlSelectDoc.SelectedItem.Value == "1" || rdlSelectDoc.SelectedItem.Value == "2" || rdlSelectDoc.SelectedItem.Value == "3")
                            {
                                txtIWContainer.Text = dt.Rows[0]["ContainerNo"].ToString();
                                txtIWPO.Text = dt.Rows[0]["PONumber"].ToString();
                                txtIWGRN.Text = dt.Rows[0]["GRNNumber"].ToString();
                            }
                            else
                            {
                                txtOWAllocation.Text = dt.Rows[0]["AllocationNo"].ToString();
                                txtOWSO.Text = dt.Rows[0]["SONumber"].ToString();
                                txtOWIssueNote.Text = dt.Rows[0]["IssueNoteNo"].ToString();
                            }
                    }
                    else
                    {
                                lblMessage.Text = "No Data Found!";
                                lblMessage.ForeColor = Color.Red;

                                txtIWContainer.Text = "";
                                txtIWPO.Text = "";
                                txtIWGRN.Text = "";

                                txtOWAllocation.Text = "";
                                txtOWIssueNote.Text = "";
                                txtOWSO.Text = "";

            }
            
        }
        #endregion cmbSearch_SelectedIndexChanged


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
                    sw.Write(",");
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
                        sw.Write(",");
                }
                sw.Write(sw.NewLine);
            }
            sw.Write(sw.NewLine);
        }

        #endregion Export To Csv

        #region FileDownload
        /// <summary>
        /// File Download
        /// </summary>

        private void FileDownload(string filename)
        {


            //string FolderPath = HttpContext.Current.Server.MapPath(".");
            //FolderPath = FolderPath + "\\Reports\\";
            //string FullFilePath = FolderPath + filename;
            FileInfo file = new FileInfo(filename);

            if (!file.Exists) return;
            if ((file.Extension == ".DAT") || (file.Extension == ".dat"))
            {
                Response.ContentType = "application/vnd.ms-excel";
                Response.AddHeader("Content-Disposition", "attachment; filename=\"" + file.Name + "\"");
                Response.AddHeader("Content-Length", file.Length.ToString());
                Response.TransmitFile(file.FullName);
                Response.Flush();
                Response.End();

            }

            if ((file.Extension == ".CSV") || (file.Extension == ".csv"))
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

        #region BindPONumber
        /// <summary>
        /// BindPONumber
        /// </summary>
        private void BindPONumber()
        {
            TatiBAL objDC = new TatiBAL();
            DataTable dt = objDC.GetPOHeader();
            CmbPONumber.DataSource = dt;
            CmbPONumber.DataMember = "PONumber";
            CmbPONumber.DataValueField = "PONumber";
           
            CmbPONumber.DataBind();

           // CmbPONumber.Items.Insert(0, "<-- Select -->");
        }


        #endregion BindPONumber

        #region BindStore
        /// <summary>
        /// BindStore
        /// </summary>
        private void BindStore()
        {
            TatiBAL objDC = new TatiBAL();
            DataTable dt = objDC.GetStore();
            CmbStoreNo.DataSource = dt;
            CmbStoreNo.DataMember = "StoreNo";
            CmbStoreNo.DataValueField = "StoreNo";
           
            CmbStoreNo.DataBind();
            CmbStoreNo.Items.Insert(0, "<-- Select -->");
        }

        #endregion BindStore

        #region BindAllocation
        /// <summary>
        /// BindAllocation
        /// </summary>
        private void BindAllocation()
        {
            TatiBAL objDC = new TatiBAL();
            objDC.Location = CmbStoreNo.Text;
            DataTable dt = objDC.GetAllocationHeader();
            ddlAllocation.DataSource = dt;
            ddlAllocation.DataMember = "AllocationNo";
            ddlAllocation.DataValueField = "AllocationNo";
            
            ddlAllocation.DataBind();
            ddlAllocation.Items.Insert(0, "<-- Select -->");
        }

        #endregion BindAllocation

        #region BindSONumber
        /// <summary>
        /// BindSONumber
        /// </summary>
        private void BindSONumber()
        {
            TatiBAL objDC = new TatiBAL();
            DataTable dt = objDC.GetAllSOHeader();
            ddlSONumber.DataSource = dt;
            ddlSONumber.DataMember = "SONumber";
            ddlSONumber.DataValueField = "SONumber";

            ddlSONumber.DataBind();
            //ddlSONumber.Items.Insert(0, "<-- Select -->");
        }

        #endregion BindSONumber


        #region BindAllPONumber
        /// <summary>
        /// BindAllPONumber
        /// </summary>
        private void BindAllPONumber()
        {
            TatiBAL objDC = new TatiBAL();
            objDC.ReportType = "1";
            objDC.DocNo = "";
            DataTable dt = objDC.GetAllPOHeader();
            ddlPONumber.DataSource = dt;
            ddlPONumber.DataMember = "PONumber";
            ddlPONumber.DataValueField = "PONumber";

            ddlPONumber.DataBind();

            // CmbPONumber.Items.Insert(0, "<-- Select -->");
        }

        #endregion BindAllPONumber

        #region BindAllContainer
        /// <summary>
        /// BindAllContainer
        /// </summary>
        private void BindAllContainer()
        {
            TatiBAL objDC = new TatiBAL();
            objDC.ReportType = "1";
            objDC.DocNo = "";
            DataTable dt = objDC.GetAllPOHeader();
            ddlContainerNo.DataSource = dt;
            ddlContainerNo.DataMember = "ContainerNo";
            ddlContainerNo.DataValueField = "ContainerNo";

            ddlContainerNo.DataBind();

            // CmbPONumber.Items.Insert(0, "<-- Select -->");
        }

        #endregion BindAllContainer

        #region BindAllPOGrn
        /// <summary>
        /// BindAllPOGrn
        /// </summary>
        private void BindAllPOGrn()
        {
            TatiBAL objDC = new TatiBAL();
            objDC.ReportType = "1";
            objDC.DocNo = "";
            DataTable dt = objDC.GetAllPOHeader();
            ddlPOGrn.DataSource = dt;
            ddlPOGrn.DataMember = "GRNNumber";
            ddlPOGrn.DataValueField = "GRNNumber";

            ddlPOGrn.DataBind();

            // CmbPONumber.Items.Insert(0, "<-- Select -->");
        }



        #endregion BindAllPOGrn

        #region BindDoc
        /// <summary>
        /// BindDoc
        /// </summary>
        private void BindDoc(string ReportType,string FieldName)
        {
            TatiBAL objDC = new TatiBAL();
            objDC.DocNo = "";
            objDC.ReportType = ReportType;
            DataTable dt = objDC.GetAllPOHeader();
            cmbSearch.DataSource = dt;
            cmbSearch.DataMember = FieldName;
            cmbSearch.DataValueField = FieldName;

            cmbSearch.DataBind();
            cmbSearch.Items.Insert(0, "<-- Select -->");
        }


        #endregion BindDoc

        #endregion Methods

      
    }
}