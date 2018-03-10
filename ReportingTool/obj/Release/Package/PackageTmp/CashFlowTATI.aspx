<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="CashFlowTATI.aspx.cs" Inherits="ReportingTool.CashFlowTATI" %>
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
    <table style="width: 100%;" border="0">
        <tr>
            <td class="text-admin-panel" width="20%">
                Cash Flow Report BTC Fashion
            </td>
            <td>
                 <asp:Label ID="lblMessage" runat="server" ForeColor="Red" ClientIDMode="Static"></asp:Label>&nbsp;
            </td>
        </tr>
    </table>

    <table>
           <tr>
                <td>Location </td>
                <td>
                    <asp:TextBox ID="txtLocation" runat="server" Text="" ></asp:TextBox>
                  
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
                <td>KSA Exchange Rate</td>
                <td>
                    <asp:TextBox ID="txtKsaRate" runat="server" Text="0.967"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator13" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtKsaRate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator4" ControlToValidate="txtKsaRate" runat="server"
                        ErrorMessage="Number Only" Display="Dynamic" ValidationExpression="[0-9]*\.?[0-9]*"></asp:RegularExpressionValidator>
                
                </td>
             </tr>

            <tr>
                <td>
                    Week No.
                </td>
                <td>
                    <asp:TextBox ID="txtWeekNo" runat="server" Text="1"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtWeekNo" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                </td>
            </tr>

             <tr>
                <td>
                   Month 
                </td>
                <td>
                     <asp:DropDownList ID="ddlMonth" runat="server">
                        <asp:ListItem Value="1" Text="Jan"></asp:ListItem>
                        <asp:ListItem Value="2" Text="Feb"></asp:ListItem>
                        <asp:ListItem Value="3" Text="Mar"></asp:ListItem>
                        <asp:ListItem Value="4" Text="Apr"></asp:ListItem>

                        <asp:ListItem Value="5" Text="May"></asp:ListItem>
                        <asp:ListItem Value="6" Text="Jun"></asp:ListItem>
                        <asp:ListItem Value="7" Text="Jul"></asp:ListItem>
                        <asp:ListItem Value="8" Text="Aug"></asp:ListItem>

                        <asp:ListItem Value="9" Text="Sep"></asp:ListItem>
                        <asp:ListItem Value="10" Text="Oct"></asp:ListItem>
                        <asp:ListItem Value="11" Text="Nov"></asp:ListItem>
                        <asp:ListItem Value="12" Text="Dec"></asp:ListItem>

                    </asp:DropDownList>

                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="ddlMonth" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                </td>
            </tr>
        
            <tr>
                <td>
                    Year
                </td>
                <td>
                    <asp:DropDownList ID="ddlYear" runat="server">
                        <asp:ListItem Value="2016" Text="2016-2017"></asp:ListItem>
                        <asp:ListItem Value="2017" Text="2017-2018"></asp:ListItem>
                    </asp:DropDownList>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator8" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="ddlYear" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                </td>
            </tr>

        <tr>
                <td>
                    Reports
                </td>
                <td>
                    
                    <asp:CheckBox ID="chkCashFlow" Text="Cash Flow TATI" Checked="false" runat="server" />
                    <asp:CheckBox ID="chkProfitLoss" Text="Profit Loss TATI" Checked="false" runat="server" />

                    <asp:CheckBox ID="chkCashFlowMY" Text="Cash Flow MY" Checked="false" runat="server" />
                    <asp:CheckBox ID="chkProfitLossMY" Text="Profit Loss MY" Checked="false" runat="server" />
                   
                </td>
            </tr>


        <tr>
            <td>&nbsp;</td>
            <td>
                <asp:Button ID="btnGenerate" runat="server" Text="Generate" style="width: 84px;" ValidationGroup="GenerateReport" ClientIDMode="Static" OnClick="btnGenerate_Click"  />
                &nbsp;<asp:Button ID="btnDownloadCashFlow" runat="server" Text="CashFlow Tati" style="width: 120px;" ValidationGroup="GenerateReport" ClientIDMode="Static" Visible="False" OnClick="btnDownloadCashFlow_Click" />

                &nbsp;<asp:Button ID="btnDownloadCashFlowMY" runat="server" Text="CashFlow MY" style="width: 120px;" ValidationGroup="GenerateReport" ClientIDMode="Static" Visible="False" OnClick="btnDownloadCashFlowMY_Click"  />

                &nbsp;<asp:Button ID="btnProfitLoss" runat="server" Text="Profit & Loss Tati" style="width: 140px;" ValidationGroup="GenerateReport" ClientIDMode="Static" Visible="False" OnClick="btnProfitLoss_Click"  />

                &nbsp;<asp:Button ID="btnProfitLossMY" runat="server" Text="Profit & Loss MY" style="width: 130px;" ValidationGroup="GenerateReport" ClientIDMode="Static" Visible="False" OnClick="btnProfitLossMY_Click"  />
            </td>
        </tr>
    </table>

</asp:Content>
