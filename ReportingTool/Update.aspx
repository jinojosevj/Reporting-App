<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Update.aspx.cs" Inherits="Test.Update" %>
<asp:Content ID="Content1" ContentPlaceHolderID="phdBodyContent" runat="server">
    
     <script type="text/javascript">
         $(document).ready(function () {
             $('#btnUpdate').click(function () {
                 $('#btnUpdate').hide();


                 $('#lblMessage').text("Table Updation Is Going On ...");
                 $('#lblMessage').css("color", "Orange");
                 $('#lblMessage').show();
             });
         });

       </script>
    <table width="100%" border="0">
        <tr>
            <td class="text-admin-panel" width="20%">
                Update Tables
            </td>
            <td>
                <asp:Label ID="lblMessage" runat="server" ForeColor="Red" ClientIDMode="Static"></asp:Label>&nbsp;
               
            </td>
        </tr>
    </table>
    
    
    <table style="width:100%;">
        <tr>
            <td style="width: 20%;">
                Item Master
            </td>
            <td>
                <asp:RadioButtonList ID="Rdltem" runat="server" RepeatDirection="Horizontal" OnSelectedIndexChanged="Rdltem_SelectedIndexChanged" AutoPostBack="True">
                        <asp:ListItem  Selected="True" Value="0">Refresh</asp:ListItem>
                        <asp:ListItem Value="1">No Refresh</asp:ListItem>
                 </asp:RadioButtonList> 
            </td>
        </tr>

        <tr id="trPromotion" runat="server">
            <td>
                &nbsp;
            </td>
            <td>
                <asp:Panel ID="panPromotion" runat="server" GroupingText="Promotions" Width="500px">
                    <div>UAE :</div> <asp:TextBox  ID="txtUaeOffer" runat="server" TextMode="MultiLine" Columns="50"></asp:TextBox>
                    <div>Bahrain :</div> <asp:TextBox ID="txtBahrainOffer" runat="server" TextMode="MultiLine" Columns="50"></asp:TextBox>
                    <div>Jordan :</div> <asp:TextBox ID="txtJordanOffer" runat="server" TextMode="MultiLine" Columns="50"></asp:TextBox>
                    <div>Oman :</div> <asp:TextBox ID="txtOmanOffer" runat="server" TextMode="MultiLine" Columns="50"></asp:TextBox>
                    <div>Qatar :</div> <asp:TextBox ID="txtQatarOffer" runat="server" TextMode="MultiLine" Columns="50"></asp:TextBox>
                    <div>KSA :</div> <asp:TextBox ID="txtKsaOffer" runat="server" TextMode="MultiLine" Columns="50"></asp:TextBox>
                </asp:Panel>

            </td>
        </tr>

               
        <tr>
            <td>
                Item Ledger Entry
            </td>
            <td>
                <asp:RadioButtonList ID="RdlItemLedger" runat="server" RepeatDirection="Horizontal">
                        <asp:ListItem Value="0">Refresh</asp:ListItem>
                        <asp:ListItem Selected="True" Value="1">Update</asp:ListItem>
                        <asp:ListItem Value="2">No Refresh</asp:ListItem>
                 </asp:RadioButtonList> 

            </td>
        </tr>

        <tr>
            <td style="height: 34px">
                Value Entry
            </td>
            <td style="height: 34px">
                <asp:RadioButtonList ID="RdlValueEntry" runat="server" RepeatDirection="Horizontal">
                        <asp:ListItem Value="0">Refresh</asp:ListItem>
                        <asp:ListItem Selected="True" Value="1">Update</asp:ListItem>
                        <asp:ListItem Value="2">No Refresh</asp:ListItem>
                 </asp:RadioButtonList>
            </td>
        </tr>
        
        
        <tr>
            <td>
                Store Footfall Register 
            </td>
            <td>
                <asp:RadioButtonList ID="RdlFootFall" runat="server" RepeatDirection="Horizontal">
                        <asp:ListItem  Value="1">Refresh</asp:ListItem>
                        <asp:ListItem Selected="True" Value="1">No Refresh</asp:ListItem>
                </asp:RadioButtonList>
            </td>
        </tr>


        <tr>
            <td>
                Transaction Header 
            </td>
            <td>
                <asp:RadioButtonList ID="RdlTransHeader" runat="server" RepeatDirection="Horizontal">
                        <asp:ListItem Selected="True" Value="0">Refresh</asp:ListItem>
                        <asp:ListItem  Value="1">No Refresh</asp:ListItem>
                </asp:RadioButtonList>
            </td>
        </tr>

        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                <asp:Button ID="btnUpdate" runat="server" Text="Update" OnClick="btnUpdate_Click"  ClientIDMode="Static"/>
            </td>
        </tr>

    </table>

</asp:Content>
