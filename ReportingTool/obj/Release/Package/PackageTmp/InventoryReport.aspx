<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="InventoryReport.aspx.cs" Inherits="ReportingTool.InventoryReport" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>
<asp:Content ID="Content1" ContentPlaceHolderID="phdBodyContent" runat="server">
    <ajaxToolkit:ToolkitScriptManager  ID="ScriptManager1" runat="server"></ajaxToolkit:ToolkitScriptManager> 
      <script type="text/javascript">
          $(document).ready(function () {
              $('#btnGenerate').click(function () {
                  $('#btnGenerate').hide();
              });
          });
       </script>

    <table style="width: 100%;" border="0">
        <tr>
            <td class="text-admin-panel" width="20%">
                Inventory Report
            </td>
            <td>
                
                 <asp:Label ID="lblMessage" runat="server" ForeColor="Red" ClientIDMode="Static"></asp:Label>&nbsp;
               
            </td>
        </tr>
    </table>
    
    <table style="width: 100%;">
        <tr>
                <td>From Date</td>
                <td>
                    
                    <asp:TextBox ID="txtFromDate" runat="server"></asp:TextBox>
                    <ajaxToolkit:CalendarExtender ID="txtFromDate_CalendarExtender" runat="server" Enabled="True" TargetControlID="txtFromDate" Format="dd/MM/yyyy" >
                    </ajaxToolkit:CalendarExtender>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" display="Dynamic" runat="server" ErrorMessage="This Field is Required" ControlToValidate="txtFromDate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td>To Date</td>
                <td>
                        <asp:TextBox ID="txtToDate" runat="server"></asp:TextBox>
                        <ajaxToolkit:CalendarExtender ID="txtToDate_CalendarExtender" runat="server" Enabled="True" TargetControlID="txtToDate" Format="dd/MM/yyyy" >
                        </ajaxToolkit:CalendarExtender>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtToDate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                        
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                    <asp:Button ID="btnGenerate" runat="server" Text="Generate" style="width: 84px;" ValidationGroup="GenerateReport" ClientIDMode="Static" OnClick="btnGenerate_Click"/>
                    &nbsp;<asp:Button ID="btnDownload" runat="server" Text="Download" style="width: 110px;"  ClientIDMode="Static" Visible="false" OnClick="btnDownload_Click"/>

                </td>
                
            </tr>

    </table>
</asp:Content>
