<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="DCReports.aspx.cs" Inherits="ReportingTool.DCReports" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>
<asp:Content ID="Content1" ContentPlaceHolderID="phdBodyContent" runat="server">
     <ajaxToolkit:ToolkitScriptManager  ID="ScriptManager1" runat="server"></ajaxToolkit:ToolkitScriptManager> 
     <script type="text/javascript">
         $(document).ready(function () {
             $('#btnGenerate').click(function () {
                 $('#btnGenerate').hide();

                 $('#lblMessage').text("Report Generation Is Going On ...");
                 $('#lblMessage').css("color", "Orange");
                 $('#lblMessage').show();
             });
         });

       </script>
     <script type="text/javascript">
          $(document).ready(function () {

              $('#mnReport').hide();
              $('#mnTati').hide();
                    
        });
       </script>

     <table style="width: 100%;" border="0">
        <tr>
            <td class="text-admin-panel" width="20%">
                DC Reports
            </td>
            <td>
                 <asp:Label ID="lblMessage" runat="server" ForeColor="Red" ClientIDMode="Static"></asp:Label>&nbsp;
            </td>
        </tr>
    </table>

     <table>
         <tr>
            <td>
            Select Report Type:
            </td>
            <td>
                   <asp:DropDownList ID="ddlType" runat="server">
                        <asp:ListItem Value="1" Text="Pack Extract"></asp:ListItem>
                        <asp:ListItem Value="2" Text="PO Linecode12"></asp:ListItem>
                        <asp:ListItem Value="3" Text="DC Ledger"></asp:ListItem>
                        <asp:ListItem Value="4" Text="DC Ledger All Items"></asp:ListItem>

                        <asp:ListItem Value="5" Text="DC Inbound & Outbound"></asp:ListItem>
                        <asp:ListItem Value="6" Text="Product Master"></asp:ListItem>
                    </asp:DropDownList>
            </td>
        </tr>

          <tr>
                <td>
                  Pack Id :
                </td>
                <td>
                      <asp:TextBox ID="txtPackId"   runat="server"></asp:TextBox>
                </td>
            </tr>
         <tr>
                <td>
                  CEVA Issue No :
                </td>
                <td>
                      <asp:TextBox ID="txtCevaNo"  runat="server"></asp:TextBox>
                </td>
            </tr>
             
           <tr>
                <td>
                   Pack Barcode:
                </td>
                <td>
                      <asp:TextBox ID="txtPackBarcode" TextMode="MultiLine" Rows="5"  Columns="30" runat="server"></asp:TextBox>
                </td>
            </tr>
          
            <tr>
                <td>
                   Linecode7:
                </td>
                <td>
                      <asp:TextBox ID="txtLinecode7" TextMode="MultiLine" Rows="5"  Columns="30" runat="server"></asp:TextBox>
                </td>
            </tr>

            <tr>
                <td>
                  PO Number :
                </td>
                <td>
                      <asp:TextBox ID="txtPONumber"  runat="server"></asp:TextBox>
                </td>
            </tr>

           <tr>
            <td>
            From Date
            </td>
            <td>
                <asp:TextBox ID="txtFromDate" TextMode="Date" runat="server"></asp:TextBox>
            </td>
        </tr>

          <tr>
            <td>
            To Date
            </td>
            <td>
                <asp:TextBox ID="txtToDate"  TextMode="Date" runat="server"></asp:TextBox>
            </td>
        </tr>
       
          <tr>
            <td>
                &nbsp;
            </td>
            <td>
                <asp:Button ID="btnGenerate" style="width: 100px;" runat="server" Text="Generate" OnClick="btnGenerate_Click" />
                &nbsp;&nbsp;
                <asp:Button ID="btnDownload" style="width: 100px;" runat="server" Text="Download" Visible="false" OnClick="btnDownload_Click"/>
            </td>
        </tr>
     </table>
</asp:Content>
