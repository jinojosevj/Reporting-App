<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="ItemInfo.aspx.cs" Inherits="ReportingTool.ItemInfo" %>
<asp:Content ID="Content1" ContentPlaceHolderID="phdBodyContent" runat="server">
    <table style="width: 100%;" border="0">
        <tr>
            <td class="text-admin-panel" width="20%">
                Item Info Report
            </td>
            <td>
                
                 <asp:Label ID="lblMessage" runat="server" ForeColor="Red" ClientIDMode="Static"></asp:Label>&nbsp;
               
            </td>
        </tr>
    </table>
    
        <table style="width: 100%;">

            <tr>
                <td>
                   Company :
                </td>
                <td>
                    <asp:DropDownList ID="ddlCompany" runat="server" OnSelectedIndexChanged="ddlCompany_SelectedIndexChanged" AutoPostBack="True">
                        <asp:ListItem Value="Select" Text="Select"></asp:ListItem>
                        <asp:ListItem Value="UAE" Text="UAE"></asp:ListItem>
                        <asp:ListItem Value="JORDAN" Text="JORDAN"></asp:ListItem>
                        <asp:ListItem Value="OMAN" Text="OMAN"></asp:ListItem>
                        <asp:ListItem Value="BAHRAIN" Text="BAHRAIN"></asp:ListItem>
                        <asp:ListItem Value="QATAR" Text="QATAR"></asp:ListItem>
                        <asp:ListItem Value="KSA" Text="KSA"></asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>

            <tr>
                <td>
                   Location :
                </td>
                <td>
                    <asp:DropDownList ID="ddlLocation" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
             <tr>
                <td>
                   Line Code :
                </td>
                <td>
                    <asp:TextBox ID="txtLineCode" runat="server" TextMode="MultiLine" Width="400px"></asp:TextBox>
                </td>
            </tr>

            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                    <asp:Button ID="btnGenerate" runat="server" Text="Generate" style="width: 84px;" ValidationGroup="GenerateReport"  ClientIDMode="Static" OnClick="btnGenerate_Click"/>
                    &nbsp;<asp:Button ID="btnDownload" runat="server" Text="Download-Item" style="width: 130px;"   ClientIDMode="Static" Visible="false" OnClick="btnDownload_Click"/>
                </td>
                
            </tr>
           
        </table>
</asp:Content>
