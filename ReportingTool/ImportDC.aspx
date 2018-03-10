<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="ImportDC.aspx.cs" Inherits="ReportingTool.ImportDC" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>
<asp:Content ID="Content1" ContentPlaceHolderID="phdBodyContent" runat="server">
  <ajaxToolkit:ToolkitScriptManager  ID="ScriptManager1" runat="server"></ajaxToolkit:ToolkitScriptManager> 
    <script type="text/javascript">
         $(window).ready(function () {
             $('#loading').hide();
         });

         $(document).ready(function () {
             $('#btnImport').click(function () {
                 $('#btnImport').hide();
                 $('#btnDwdLog').hide();
                 $('#btnSaveStockLedger').hide();
                
                 $('#lblMessage').text("Process Is Going On ...");
                 $('#lblMessage').css("color", "Orange");
                 $('#lblMessage').show();

                 $('#loading').show();

             });

             $('#btnSaveStockLedger').click(function () {
                 $('#btnSaveStockLedger').hide();
                 $('#lblMessage').text("Process Is Going On ...");
                 $('#lblMessage').css("color", "Orange");
                 $('#lblMessage').show();

                 $('#loading').show();

             });
         });
       </script>

     <script type="text/javascript">
          $(document).ready(function () {

              $('#mnReport').hide();
              $('#mnTati').hide();
        });
       </script>
     <div  id="loading" ></div>
     <table style="width: 100%;" border="0">
        <tr>
            <td class="text-admin-panel" width="20%">
                 Import DC Masters
            </td>
            <td>
                 <asp:Label ID="lblMessage" runat="server" ForeColor="Red" ClientIDMode="Static"></asp:Label>&nbsp;
                 
            </td>
        </tr>
    </table>
      <table>
         <tr>
            <td>
            Select Import Type:
            </td>
            <td>
                   <asp:DropDownList ID="ddlType" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlType_SelectedIndexChanged">
                        <asp:ListItem Value="1" Text="Pack Extract"></asp:ListItem>
                        <asp:ListItem Value="2" Text="Container Extract"></asp:ListItem>
                        <asp:ListItem Value="3" Text="PO GRN"></asp:ListItem>
                        <asp:ListItem Value="4" Text="Allocation"></asp:ListItem>
                        
                        <asp:ListItem Value="5" Text="SO Issue Note"></asp:ListItem>
                        <asp:ListItem Value="6" Text="Product Group Listing For Pack"></asp:ListItem>
                        <asp:ListItem Value="7" Text="Family Listing For Single"></asp:ListItem>
                        <asp:ListItem Value="8" Text="Create SO From Posted Purchse Order"></asp:ListItem>
                        
                        <asp:ListItem Value="9" Text="Delete PO"></asp:ListItem>
                        <asp:ListItem Value="10" Text="Delete SO"></asp:ListItem>
                        <asp:ListItem Value="11" Text="Delete Allocation"></asp:ListItem>
                        <asp:ListItem Value="12" Text="Extra Container Extract"></asp:ListItem>

                        <asp:ListItem Value="13" Text="Adjustment"></asp:ListItem>
                        <asp:ListItem Value="14" Text="HS Codes"></asp:ListItem>

                    </asp:DropDownList>
            </td>
        </tr>
          <tr id="trPONumber" runat="server" visible="false">
              <td>PO Number:</td>
              <td>
                  <asp:TextBox ID="txtPONumber"  runat="server" Visible="false"></asp:TextBox>

                  <ajaxToolkit:ComboBox ID="CmbPONumber" ClientIDMode="Static" runat="server" 
                    AutoPostBack="False" 
                    DropDownStyle="DropDown" 
                    AutoCompleteMode="SuggestAppend" 
                    CaseSensitive="False" 
                    CssClass="" 
                    ItemInsertLocation="Append"
                       >
                  </ajaxToolkit:ComboBox>  
                  </td>
          </tr>
           <tr id="trContainerReference" runat="server" visible="false">
              <td>Container Reference:</td>
              <td>
                    <asp:TextBox ID="txtContainerReference"  runat="server"></asp:TextBox>
              </td>
          </tr>

           <tr id="trGRNNo" runat="server" visible="false">
              <td>GRN No:</td>
              <td>
                    <asp:TextBox ID="txtGRNNo"  runat="server"></asp:TextBox>
              </td>
          </tr>
          
          <tr id="trAsOfDate" runat="server" visible="false">
              <td>As Of Date:</td>
              <td>
                    <asp:TextBox ID="txtAsOfDate" runat="server"></asp:TextBox>
                    <ajaxToolkit:CalendarExtender ID="CalendarExtender5" runat="server" Enabled="True" TargetControlID="txtAsOfDate" Format="dd/MM/yyyy" >
                    </ajaxToolkit:CalendarExtender>
              </td>
          </tr>

           <tr id="trtxtSONumber" runat="server" visible="false">
              <td>SO Number:</td>
              <td>
                    <asp:TextBox ID="txtSONumber"  runat="server"></asp:TextBox>
              </td>
          </tr>
           <tr id="trCompany" runat="server" visible="false">
              <td>Company:</td>
              <td>
                     <asp:DropDownList ID="ddlCompany" runat="server" AutoPostBack="true" >
                        <asp:ListItem Value="MATALAN-UAE HO" Text="MATALAN-UAE HO"></asp:ListItem>
                        <asp:ListItem Value="MATALAN-JORDAN HO" Text="MATALAN-JORDAN HO"></asp:ListItem>
                        <asp:ListItem Value="MATALAN-OMAN HO" Text="MATALAN-OMAN HO"></asp:ListItem>
                        <asp:ListItem Value="MATALAN-BAHRAIN HO" Text="MATALAN-BAHRAIN HO"></asp:ListItem>
                         
                        <asp:ListItem Value="MATALAN-QATAR HO" Text="MATALAN-QATAR HO"></asp:ListItem>
                        <asp:ListItem Value="MATALAN-KSA HO" Text="MATALAN-KSA HO"></asp:ListItem>
                        <asp:ListItem Value="MATALAN DC-JAFZA" Text="MATALAN DC-JAFZA"></asp:ListItem>
                    </asp:DropDownList>
              </td>
          </tr>
           <tr id="trIssueNoteNo" runat="server" visible="false">
              <td>Issue Note No.:</td>
              <td>
                    <asp:TextBox ID="txtIssueNoteNo"  runat="server"></asp:TextBox>
              </td>
          </tr>
          <tr id="trAllocationNo" runat="server" visible="false">
              <td>Allocation No:</td>
              <td>
                    <asp:TextBox ID="txtAllocationNo"  runat="server"></asp:TextBox>
              </td>
          </tr>
           <tr id="trStoreNo" runat="server" visible="false">
              <td>Store No:</td>
              <td>
                    <asp:DropDownList ID="cmbStoreNo" runat="server" AutoPostBack="true" >
                       
                    </asp:DropDownList>
              </td>
           </tr>

           <tr id="trSONumber" runat="server" visible="false">
              <td>SO Number:</td>
              <td>
                    <ajaxToolkit:ComboBox ID="cmbSONumber" ClientIDMode="Static" runat="server" 
                    AutoPostBack="False" 
                    DropDownStyle="DropDown" 
                    AutoCompleteMode="SuggestAppend" 
                    CaseSensitive="False" 
                    CssClass="" 
                    ItemInsertLocation="Append"
                       >
                  </ajaxToolkit:ComboBox>  
              </td>
           </tr>
          
             <tr id="trSelectAdjustment" runat="server" visible="false">
                  <td>
                      Select Adjustment Type:
                  </td>
                  <td>
                      <asp:RadioButtonList ID="rdlSelectDoc"  runat="server" AutoPostBack="true" RepeatDirection="Horizontal" OnSelectedIndexChanged="rdlSelectDoc_SelectedIndexChanged" >
                          <asp:ListItem Value="1" Selected="True">PO</asp:ListItem>
                          <asp:ListItem Value="2">SO</asp:ListItem>
                          <asp:ListItem Value="3">PI Adjustment</asp:ListItem>
                      </asp:RadioButtonList>

                  </td>
            </tr>

            <tr id="trDocNo" runat="server" visible="false">
              <td>Original Doc Number:</td>
                 
              <td>
                    <ajaxToolkit:ComboBox ID="cmbDocNumber" ClientIDMode="Static" runat="server" 
                    AutoPostBack="False" 
                    DropDownStyle="DropDown" 
                    AutoCompleteMode="SuggestAppend" 
                    CaseSensitive="False" 
                    CssClass="" 
                    ItemInsertLocation="Append"
                       >
                  </ajaxToolkit:ComboBox>  

              </td>
           </tr>
          <tr id="trAdjustment" runat="server" visible="false">
              <td>
                  Adjustment No.
              </td>
              <td>
                  <asp:TextBox ID="txtAdjustmentNo"  runat="server" Visible="true"></asp:TextBox>
              </td>
          </tr>

            <tr id="trAdjustmentDate" runat="server" visible="false">
              <td>
                  Adjustment Date
              </td>
              <td>
                  <asp:TextBox ID="txtAdjustmentDate" TextMode="Date"  runat="server" ></asp:TextBox>
              </td>
          </tr>

            <tr>
                <td>
                   <asp:FileUpload ID="fileuploadExcel"  runat="server" />
                </td>
            </tr>

          <tr>
            <td>
                  <asp:Button ID="btnImport" ClientIDMode="Static" style="width: 100px;" runat="server" Text="Import" OnClick="btnImport_Click"  />
                  &nbsp;&nbsp;
                  <asp:Button ID="btnDwdLog" ClientIDMode="Static" style="width: 100px;" runat="server" Text="DwdLog" Visible="false" OnClick="DwdLog_Click"   />
                  &nbsp;&nbsp;
                  <asp:Button ID="btnSaveStockLedger" ClientIDMode="Static" style="width: 150px;" runat="server" Text="SaveStockLedger" Visible="false" OnClick="btnSaveStockLedger_Click"  />
            </td>
        </tr>
      </table>
</asp:Content>
