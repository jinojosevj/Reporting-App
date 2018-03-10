<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="ExportDC.aspx.cs" Inherits="ReportingTool.ExportDC" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>
<asp:Content ID="Content1" ContentPlaceHolderID="phdBodyContent" runat="server">
    <ajaxToolkit:ToolkitScriptManager  ID="ScriptManager1" runat="server"></ajaxToolkit:ToolkitScriptManager> 

    
      <script type="text/javascript">
          $(window).ready(function () {
              $('#loading').hide();
          });

         $(document).ready(function () {
             $('#btnExport').click(function () {
                 $('#btnExport').hide();
                 $('#btnDownload').hide();
                 
                 $('#lblMessage').text("Report Generation Is Going On ...");
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
                 Export DC 
            </td>
            <td>
                 <asp:Label ID="lblMessage" runat="server" ForeColor="Red" ClientIDMode="Static"></asp:Label>&nbsp;
            </td>
        </tr>
    </table>
      <table>
         <tr>
            <td>
            Select Export Type:
            </td>
            <td>
                 <asp:DropDownList ID="ddlType" runat="server" AutoPostBack="true"  OnSelectedIndexChanged="ddlType_SelectedIndexChanged">
                        <asp:ListItem Value="1" Text="PO"></asp:ListItem>
                        <asp:ListItem Value="2" Text="PO Item Master Pack"></asp:ListItem>
                        <asp:ListItem Value="3" Text="PO Item Master Single/Inner"></asp:ListItem>
                        <asp:ListItem Value="4" Text="Create SO"></asp:ListItem>
                     
                        <asp:ListItem Value="5" Text="Product Code Listing"></asp:ListItem>
                        <asp:ListItem Value="6" Text="Item Family Listing"></asp:ListItem>
                        <asp:ListItem Value="7" Text="Stock Ledger By Linecode7"></asp:ListItem>
                        <asp:ListItem Value="8" Text="Stock Ledger By Pack Barcode"></asp:ListItem>
                     
                        <asp:ListItem Value="9" Text="Pack Extract By Linecode7"></asp:ListItem>
                        <asp:ListItem Value="10" Text="Pack Extract By Pack Barcode"></asp:ListItem>
                        <asp:ListItem Value="11" Text="Issue Note LineCode12"></asp:ListItem>
                        <asp:ListItem Value="12" Text="Export SO"></asp:ListItem>
                     
                        <asp:ListItem Value="13" Text="SO Header"></asp:ListItem>
                        <asp:ListItem Value="14" Text="Inward"></asp:ListItem>
                        <asp:ListItem Value="17" Text="Inward/Outward"></asp:ListItem>
                        <asp:ListItem Value="20" Text="Pack Extract"></asp:ListItem>

                        <asp:ListItem Value="21" Text="Allocation By Packbarcode"></asp:ListItem>
                        <asp:ListItem Value="22" Text="Allocation By Linecode7/12"></asp:ListItem>
                        <asp:ListItem Value="23" Text="SO By Packbarcode"></asp:ListItem>
                        <asp:ListItem Value="24" Text="SO By Linecode7/12"></asp:ListItem>
                        <asp:ListItem Value="25" Text="Adjustment"></asp:ListItem>
                         
                        <asp:ListItem Value="26" Text="HS Codes"></asp:ListItem>
                 </asp:DropDownList>
            </td>
        </tr>
          <tr id="trPONumber" runat="server">
              <td >PO Number:</td>
              <td>
                    <ajaxToolkit:ComboBox ID="CmbPONumber" ClientIDMode="Static" runat="server" 
                    AutoPostBack="false" 
                    DropDownStyle="DropDown" 
                    AutoCompleteMode="SuggestAppend" 
                    CaseSensitive="False" 
                    CssClass="" 
                    ItemInsertLocation="Append"
                       >
                  </ajaxToolkit:ComboBox>
                  
              </td>
          </tr>
           <tr id ="trSONumber" runat="server" visible="false">
              <td>SO Number:</td>
              <td>
                    <asp:TextBox ID="txtSONumber" runat="server"></asp:TextBox>
              </td>
          </tr>
          <tr id ="trCmbSONumber" runat="server" visible="false" >
              <td>SO Number:</td>
              <td>
                    <asp:DropDownList ID="ddlSONumber" runat="server" AutoPostBack="true" >
                   </asp:DropDownList>
              </td>
          </tr>
         <tr id="trStoreNo" runat="server" visible="false">
             <td>Store No:</td>
             <td>
                <ajaxToolkit:ComboBox ID="CmbStoreNo" ClientIDMode="Static" runat="server" 
                    AutoPostBack="true" 
                    DropDownStyle="DropDown" 
                    AutoCompleteMode="SuggestAppend" 
                    CaseSensitive="False" 
                    CssClass="" 
                    ItemInsertLocation="Append" OnSelectedIndexChanged="CmbStoreNo_SelectedIndexChanged"
                       >
                  </ajaxToolkit:ComboBox>  
             </td>

         </tr>
          <tr>
              <td></td>
               <td>
                   <asp:DropDownList ID="ddlAllocation" Visible="false"  runat="server" AutoPostBack="true" OnTextChanged="ddl_TextChanged"  >
                   </asp:DropDownList>
                  </td>
          </tr>
           <tr id="trAllocationNo" runat="server" visible="false">
              <td>AllocationNo:</td>
               <td>
                    <asp:TextBox ReadOnly="true" ID="txtAllocationNo" TextMode="MultiLine" Rows="5"  Columns="30" runat="server"></asp:TextBox>
              </td>
           </tr>
           <tr id="trLineCode7" runat="server" visible="false">
              <td>LineCode7:</td>
               <td>
                    <asp:TextBox  ID="txtLineCode7" TextMode="MultiLine" Rows="5"  Columns="30" runat="server"></asp:TextBox>
              </td>
           </tr>
          <tr id="trPackBarcode" runat="server" visible="false">
              <td>PackBarcode:</td>
               <td>
                    <asp:TextBox  ID="txtPackBarcode" TextMode="MultiLine" Rows="5"  Columns="30" runat="server"></asp:TextBox>
              </td>
           </tr>
          <tr id="trIssueNo" runat="server" visible="false">
              <td>Issue No:</td>
               <td>
                    <asp:TextBox  ID="txtIssueNo"  runat="server"></asp:TextBox>
              </td>
           </tr>



          <tr id="trContainerNo" runat="server" visible="false">
              <td>Container No:</td>
               <td>
                    <asp:DropDownList ID="ddlContainerNo"   runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlContainerNo_SelectedIndexChanged" >
                   </asp:DropDownList>
              </td>
          </tr>

         

          <tr id="trddlPONumber" runat="server" visible="false">
              <td>PO Number:</td>
               <td>
                    <asp:DropDownList ID="ddlPONumber"   runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlPONumber_SelectedIndexChanged" >
                   </asp:DropDownList>
              </td>
          </tr>
          


          <tr id="trPOGrn" runat="server" visible="false">
              <td>PO GRN No:</td>
               <td>
                    <asp:DropDownList ID="ddlPOGrn"  runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlPOGrn_SelectedIndexChanged" >
                   </asp:DropDownList>
              </td>
          </tr>

          
         
           <tr id="trSelectDoc" runat="server" visible="false">
              <td>
                  Select Document:
              </td>
              <td>
                  <asp:RadioButtonList ID="rdlSelectDoc"  runat="server" AutoPostBack="true" RepeatDirection="Horizontal" OnSelectedIndexChanged="rdlSelectDoc_SelectedIndexChanged" >
                      <asp:ListItem Value="1" Selected="True">Container</asp:ListItem>
                      <asp:ListItem Value="2">PO</asp:ListItem>
                      <asp:ListItem Value="3">GRN</asp:ListItem>
                      <asp:ListItem Value="4">Allo.</asp:ListItem>
                      <asp:ListItem Value="5">SO</asp:ListItem>
                      <asp:ListItem Value="6">Issue</asp:ListItem>
                  </asp:RadioButtonList>
              </td>

          </tr>
          <tr id="trSearch" runat="server" visible="false">
             <td>Search:</td>
             <td>
                <ajaxToolkit:ComboBox ID="cmbSearch" ClientIDMode="Static" runat="server" 
                    AutoPostBack="true" 
                    DropDownStyle="DropDown" 
                    AutoCompleteMode="SuggestAppend" 
                    CaseSensitive="False" 
                    CssClass="" 
                    ItemInsertLocation="Append" OnSelectedIndexChanged="cmbSearch_SelectedIndexChanged" 
                       >
                  </ajaxToolkit:ComboBox>  
             </td>

         </tr>
         
          <tr id="trOutwardDocs" runat="server" visible="false">
              <td>Allocation:</td>
               <td>
                    <asp:TextBox ReadOnly="true" ID="txtOWAllocation"  runat="server"></asp:TextBox>
              </td>
              <td>SO:</td>
               <td>
                    <asp:TextBox ReadOnly="true" ID="txtOWSO"  runat="server"></asp:TextBox>
              </td>
              <td>Issue:</td>
               <td>
                    <asp:TextBox ReadOnly="true" ID="txtOWIssueNote"  runat="server"></asp:TextBox>
              </td>
           </tr>

          <tr id="trInwardDocs" runat="server" visible="false">
              <td>Container:</td>
               <td>
                    <asp:TextBox ReadOnly="true" ID="txtIWContainer"  runat="server"></asp:TextBox>
              </td>
              <td>PO:</td>
               <td>
                    <asp:TextBox ReadOnly="true" ID="txtIWPO"  runat="server"></asp:TextBox>
              </td>
              <td>GRN:</td>
               <td>
                    <asp:TextBox ReadOnly="true" ID="txtIWGRN"  runat="server"></asp:TextBox>
              </td>
           </tr>

          <tr id="trInwardSelectFile" runat="server" visible="false">
              <td>
                  Select File:
              </td>
              <td>
                  <asp:RadioButtonList ID="rblInwardSelectFile"  runat="server" RepeatDirection="Horizontal" >
                      <asp:ListItem Value="1" Selected="True">Container</asp:ListItem>
                      <asp:ListItem Value="2">PO</asp:ListItem>
                      <asp:ListItem Value="3">PO GRN</asp:ListItem>
                  </asp:RadioButtonList>
              </td>

          </tr>

          <tr id="trRdlOWSelectFile" runat="server" visible="false">
              <td>
                  Select File:
              </td>
              <td>
                  <asp:RadioButtonList ID="RdlOWSelectFile"  runat="server" RepeatDirection="Horizontal" >
                      <asp:ListItem Value="1" Selected="True">Allocation</asp:ListItem>
                      <asp:ListItem Value="2">SO</asp:ListItem>
                      <asp:ListItem Value="3">Issue Note</asp:ListItem>
                  </asp:RadioButtonList>
              </td>

          </tr>

          <tr id="trDocumentNo" runat="server" visible="false">
              <td>Document No:</td>
               <td>
                    <asp:TextBox  ID="txtDocumentNo"  runat="server"></asp:TextBox>
              </td>
           </tr>

          <tr>
            <td>
              
            </td>
            <td>
                  <asp:Button ID="btnExport" ClientIDMode="Static" style="width: 100px;" runat="server" Text="Export" OnClick="btnExport_Click"/>
                  &nbsp;&nbsp;
                  <asp:Button ID="btnDownload" ClientIDMode="Static" style="width: 100px;" runat="server" Text="Download" Visible="false" OnClick="btnDownload_Click"/>

            </td>
        </tr>
      </table>
</asp:Content>
