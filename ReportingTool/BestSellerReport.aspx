<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="BestSellerReport.aspx.cs" Inherits="ReportingTool.BestSeller" %>
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
                Best Seller Report
            </td>
            <td>
                 <asp:Label ID="lblMessage" runat="server" ForeColor="Red" ClientIDMode="Static"></asp:Label>&nbsp;
            </td>
        </tr>
    </table>

    <table>
         <tr>
                <td>
                    Best Seller Table
                </td>
                <td>
                     <asp:RadioButtonList ID="rdlBestSeller" runat="server" RepeatDirection="Horizontal">
                         <asp:ListItem Value="1" Selected="True">Refresh</asp:ListItem>
                         <asp:ListItem Value="0">No Refresh</asp:ListItem>
                     </asp:RadioButtonList>
                </td>
            </tr>
        <tr>
            <td>Division Code :</td>
            <td>
                <asp:DropDownList ID="ddlDivisionCode" runat="server">
                        <asp:ListItem Value="All" Text="All"></asp:ListItem>
                        <asp:ListItem Value="M" Text="Menswear"></asp:ListItem>
                        <asp:ListItem Value="L" Text="Ladieswear"></asp:ListItem>
                        <asp:ListItem Value="C" Text="Childrenswear"></asp:ListItem>
                        <asp:ListItem Value="F" Text="Footwear"></asp:ListItem>
                        <asp:ListItem Value="H" Text="Homewear"></asp:ListItem>
                        <asp:ListItem Value="S" Text="Essentials"></asp:ListItem>
                </asp:DropDownList>
            </td>
        </tr>
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
                <td>UAE Exchange Rate</td>
                <td>
                    <asp:TextBox ID="txtUaeRate" runat="server" Text="0.985" ></asp:TextBox>
                     <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtUaeRate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                    <asp:RegularExpressionValidator ID="rgx" ControlToValidate="txtUaeRate" runat="server"
                        ErrorMessage="Number Only" Display="Dynamic" ValidationExpression="[0-9]*\.?[0-9]*"></asp:RegularExpressionValidator>
                </td>
            </tr>

            <tr>
                <td>Jordan Exchange Rate</td>
                <td>
                    <asp:TextBox ID="txtJordanRate" runat="server" Text="5.09"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtJordanRate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                     <asp:RegularExpressionValidator ID="RegularExpressionValidator1" ControlToValidate="txtJordanRate" runat="server"
                        ErrorMessage="Number Only" Display="Dynamic" ValidationExpression="[0-9]*\.?[0-9]*"></asp:RegularExpressionValidator>
                
                </td>
            </tr>

            <tr>
                <td>Oman Exchange Rate</td>
                <td>
                    <asp:TextBox ID="txtOmanRate" runat="server" Text="9.4" ></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator5" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtOmanRate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator2" ControlToValidate="txtOmanRate" runat="server"
                        ErrorMessage="Number Only" Display="Dynamic" ValidationExpression="[0-9]*\.?[0-9]*"></asp:RegularExpressionValidator>
                
                </td>
            </tr>

            <tr>
                <td>Bahrain Exchange Rate</td>
                <td>
                    <asp:TextBox ID="txtBahrainRate" runat="server" Text="9.61"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator6" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtBahrainRate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator3" ControlToValidate="txtBahrainRate" runat="server"
                        ErrorMessage="Number Only" Display="Dynamic" ValidationExpression="[0-9]*\.?[0-9]*"></asp:RegularExpressionValidator>
                
                </td>
            </tr>
            
        <tr>
                <td>Ksa Exchange Rate</td>
                <td>
                    <asp:TextBox ID="txtKsaRate" runat="server" Text="0.967"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator8" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtKsaRate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator4" ControlToValidate="txtKsaRate" runat="server"
                        ErrorMessage="Number Only" Display="Dynamic" ValidationExpression="[0-9]*\.?[0-9]*"></asp:RegularExpressionValidator>
                
                </td>
            </tr>

            <tr>
                <td>
                    Week No
                </td>
                <td>
                    <asp:TextBox ID="txtWeekNo" runat="server" TextMode="Number" MaxLength="2"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator7" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtWeekNo" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                </td>
            </tr>

        <tr>
            <td>&nbsp;</td>
            <td>
                <asp:Button ID="btnGenerate" runat="server" Text="Generate" style="width: 84px;" ValidationGroup="GenerateReport" ClientIDMode="Static" OnClick="btnGenerate_Click" />
                &nbsp;<asp:Button ID="btnDownload" runat="server" Text="BestSeller" style="width: 94px;" ValidationGroup="GenerateReport" ClientIDMode="Static" OnClick="btnDownload_Click" Visible="False"/>
                &nbsp;<asp:Button ID="btnBestSellerLC7" runat="server" Text="BestSellerLC7" style="width: 115px;" ValidationGroup="GenerateReport" ClientIDMode="Static" Visible="False" OnClick="btnBestSellerLC7_Click"/>
            </td>
        </tr>
    </table>
</asp:Content>
