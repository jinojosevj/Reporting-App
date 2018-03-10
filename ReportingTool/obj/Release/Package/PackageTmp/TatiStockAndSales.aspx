<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="TatiStockAndSales.aspx.cs" Inherits="ReportingTool.TatiStockAndSales" %>
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
                TATI Stock And Sales
            </td>
            <td>
                 <asp:Label ID="lblMessage" runat="server" ForeColor="Red" ClientIDMode="Static"></asp:Label>&nbsp;
            </td>
        </tr>
    </table>

    <table>
        
         <tr>
                <td>Posting Date :</td>
                <td>
                    
                    <asp:TextBox ID="txtPostingDate" runat="server"></asp:TextBox>
                    <ajaxToolkit:CalendarExtender ID="txtPostingDate_CalendarExtender" runat="server" Enabled="True" TargetControlID="txtPostingDate" Format="dd/MM/yyyy" >
                    </ajaxToolkit:CalendarExtender>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" display="Dynamic" runat="server" ErrorMessage="This Field is Required" ControlToValidate="txtPostingDate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                </td>
            </tr>
        
     
        <tr>
            <td>&nbsp;</td>
            <td>
                <asp:Button ID="btnGenerate" runat="server" Text="Generate-Sales" style="width: 150px;" ValidationGroup="GenerateReport" ClientIDMode="Static" OnClick="btnGenerate_Click" />
                &nbsp;<asp:Button ID="btnGenerateStock" runat="server" Text="Generate-Stock" style="width: 150px;" ValidationGroup="GenerateReport" ClientIDMode="Static" OnClick="btnGenerateStock_Click" />
               
            </td>
        </tr>

        <tr>
            <td>&nbsp;</td>
            <td>
                 &nbsp;<asp:Button ID="btnSales" runat="server" Text="Dwd-Sales-4728" style="width: 120px;" ValidationGroup="GenerateReport" ClientIDMode="Static"  Visible="False" OnClick="btnSales_Click"/>
                 &nbsp;<asp:Button ID="btnDwdSales4729" runat="server" Text="Dwd-Sales-4729" style="width: 120px;" ValidationGroup="GenerateReport" ClientIDMode="Static"  Visible="False" OnClick="btnDwdSales4729_Click"/>
                 &nbsp;<asp:Button ID="btnDwdSales4731" runat="server" Text="Dwd-Sales-4731" style="width: 120px;" ValidationGroup="GenerateReport" ClientIDMode="Static"  Visible="False" OnClick="btnDwdSales4731_Click"/>

                 &nbsp;<asp:Button ID="btnStock" runat="server" Text="Dwd-Stock-4728" style="width: 120px;" ValidationGroup="GenerateReport" ClientIDMode="Static" Visible="False" OnClick="btnStock_Click" />
                 &nbsp;<asp:Button ID="btnDwdStock4729" runat="server" Text="Dwd-Stock-4729" style="width: 120px;" ValidationGroup="GenerateReport" ClientIDMode="Static" Visible="False" OnClick="btnDwdStock4729_Click" />
                 &nbsp;<asp:Button ID="btnDwdStock4731" runat="server" Text="Dwd-Stock-4731" style="width: 120px;" ValidationGroup="GenerateReport" ClientIDMode="Static" Visible="False" OnClick="btnDwdStock4731_Click" />
            </td>

        </tr>
    </table>
</asp:Content>
