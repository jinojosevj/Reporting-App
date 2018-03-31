<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="FinanceReports.aspx.cs" Inherits="ReportingTool.FinanceReports" %>
<asp:Content ID="Content1" ContentPlaceHolderID="phdBodyContent" runat="server">
     <script type="text/javascript">
         $(window).ready(function () {
             $('#loading').hide();
         });

         $(document).ready(function () {
             $('#btnGenerate').click(function () {
                 $('#btnGenerate').hide();
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
              $('#mnDCApp').hide();
              $('#mnDCStock').hide();
              
        });
       </script>
     <div  id="loading" ></div>
    <table style="width: 100%;" border="0">
        <tr>
            <td class="text-admin-panel" width="20%">
                Finance Reports
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
                        <asp:ListItem Value="1" Text="Inventory Listing"></asp:ListItem>
                        <asp:ListItem Value="2" Text="Journal Entry Testing"></asp:ListItem>
                        <asp:ListItem Value="3" Text="Weighted Average Testing"></asp:ListItem>
                        <asp:ListItem Value="4" Text="Sales"></asp:ListItem>
                        <asp:ListItem Value="5" Text="Stock Cost"></asp:ListItem>

                    </asp:DropDownList>
            </td>
        </tr>

         <tr>
            <td>
                Country:
            </td>
            <td>
                   <asp:DropDownList ID="ddlCountry" runat="server"  AutoPostBack="True">
                        <asp:ListItem Value="Select" Text="Select"></asp:ListItem>
                        <asp:ListItem Value="UAE" Text="UAE"></asp:ListItem>
                        <asp:ListItem Value="JOR" Text="JORDAN"></asp:ListItem>
                        <asp:ListItem Value="OMAN" Text="OMAN"></asp:ListItem>
                        <asp:ListItem Value="BAH" Text="BAHRAIN"></asp:ListItem>
                        <asp:ListItem Value="QAT" Text="QATAR"></asp:ListItem>
                        <asp:ListItem Value="KSA" Text="KSA"></asp:ListItem>
                        <asp:ListItem Value="DC" Text="DC"></asp:ListItem>
                    </asp:DropDownList>
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
                <asp:Button ID="btnGenerate" ClientIDMode="Static" style="width: 100px;" runat="server" Text="Generate" OnClick="btnGenerate_Click"  />
                &nbsp;&nbsp;
                <asp:Button ID="btnDownload" ClientIDMode="Static" style="width: 100px;" runat="server" Text="Download" Visible="false" OnClick="btnDownload_Click"  />
            </td>
        </tr>
     </table>
</asp:Content>
