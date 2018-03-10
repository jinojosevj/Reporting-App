<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="TATIMISReports.aspx.cs" Inherits="ReportingTool.TATIMISReports" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>
<asp:Content ID="Content1" ContentPlaceHolderID="phdBodyContent" runat="server">
     <ajaxToolkit:ToolkitScriptManager  ID="ScriptManager1" runat="server"></ajaxToolkit:ToolkitScriptManager> 
     <script type="text/javascript">
         $(document).ready(function () {
             $('#btnGenerate').click(function () {
                 $('#btnGenerate').hide();
                 $('#btnDownload').hide();
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
                TATI MIS Reports
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
                        <asp:ListItem Value="1" Text="Stock Status"></asp:ListItem>
                        <asp:ListItem Value="2" Text="Item Info"></asp:ListItem>
                        <asp:ListItem Value="3" Text="Inventory Adjustment"></asp:ListItem>
                        <asp:ListItem Value="4" Text="Update Offer Price"></asp:ListItem>

                        <asp:ListItem Value="5" Text="BestSeller Report By Qty"></asp:ListItem>
                        <asp:ListItem Value="6" Text="BestSeller Report By Amount"></asp:ListItem>
                    </asp:DropDownList>
            </td>
        </tr>

         <tr>
            <td>
                Country:
            </td>
            <td>
                   <asp:DropDownList ID="ddlCountry" runat="server" AutoPostBack="True">
                        <asp:ListItem Value="JOR" Text="Jordan"></asp:ListItem>
                    </asp:DropDownList>
            </td>
        </tr>
          <tr>
                <td style="height: 32px">
                   Location :
                </td>
                <td style="height: 32px">
                    <asp:DropDownList ID="ddlLocation" runat="server" >
                        <asp:ListItem Value="4728" Text="4728"></asp:ListItem>
                        <asp:ListItem Value="4729" Text="4729"></asp:ListItem>
                        <asp:ListItem Value="4731" Text="4731"></asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>

         <tr>
            <td>
                Promotion Numbers
            </td>
            <td>
                <asp:TextBox ID="txtPromoNumbers" TextMode="MultiLine" Columns="50"  runat="server"></asp:TextBox>
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
                <asp:Button ID="btnGenerate" style="width: 100px;" ClientIDMode="Static" runat="server" Text="Generate" OnClick="btnGenerate_Click" />
                &nbsp;&nbsp;
                <asp:Button ID="btnDownload" style="width: 100px;" ClientIDMode="Static" runat="server" Text="Download" Visible="false" OnClick="btnDownload_Click"  />
            </td>
        </tr>
     </table>

</asp:Content>
