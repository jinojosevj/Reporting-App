<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="DcStockReport.aspx.cs" Inherits="ReportingTool.DcStockReport" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>

<asp:Content ID="Content1" ContentPlaceHolderID="phdBodyContent" runat="server">
     <ajaxtoolkit:toolkitscriptmanager  ID="ScriptManager1" runat="server"></ajaxtoolkit:toolkitscriptmanager>
    
      <script type="text/javascript">
          $(document).ready(function () {

              $('#liUpdate').hide();
              $('#liStockStatus').hide();
              $('#liVisitor').hide();

              $('#btnGenerate').click(function () {
                  $('#btnGenerate').hide();
              });
          });
       </script>
    
    <table width="100%" border="0">
        <tr>
            <td class="text-admin-panel" width="20%">
                Stock Summary Report
            </td>
            <td>
                
                 <asp:Label ID="lblMessage" runat="server" ForeColor="Red" ClientIDMode="Static"></asp:Label>&nbsp;
               
            </td>
        </tr>
    </table>
    
    <table  style="width:100%; border:0">
        <tr>
            <td>Loction Code :</td>
            <td><asp:TextBox ID="txtLocation" runat="server"></asp:TextBox></td>
        </tr>
        <tr>
            <td>As Of Date:</td>
            <td>
                    <asp:TextBox ID="txtAsofDate" runat="server"></asp:TextBox>
                    <ajaxtoolkit:calendarextender ID="txtAsofDate_CalendarExtender" runat="server" Enabled="True" TargetControlID="txtAsofDate" Format="dd/MM/yyyy" >
                    </ajaxtoolkit:calendarextender>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" display="Dynamic" runat="server" ErrorMessage="This Field is Required" ControlToValidate="txtAsofDate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>

            </td>
        </tr>
        <tr>
            <td>&nbsp;</td>
            <td>
                    <asp:Button ID="btnGenerate" runat="server" Text="Generate" style="width: 84px;" ValidationGroup="GenerateReport"  ClientIDMode="Static" OnClick="btnGenerate_Click"/>
                    
                   &nbsp; <asp:Button ID="btnDownload" runat="server" Text="Stock Summary Report" style="width: 165px;"  Visible="false" ClientIDMode="Static" />

            </td>
        </tr>
    </table>
</asp:Content>
