<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="MatalanSalesAndStock.aspx.cs" Inherits="ReportingTool.MatalanSalesAndStock" %>
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
                Matalan Sales And Stock Files
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
                        <asp:ListItem Value="1" Text="Sales"></asp:ListItem>
                        <asp:ListItem Value="2" Text="Stock"></asp:ListItem>
                    </asp:DropDownList>
            </td>
        </tr>
          <tr>
            <td>
            Sequence No. Start
            </td>
            <td>
                <asp:TextBox ID="txtSequenceNo"  runat="server"></asp:TextBox>
            </td>
          </tr>
          
           <tr>
            <td>
             Store No:
            </td>
            <td>
                <asp:TextBox ID="txtStoreNo"  runat="server"></asp:TextBox>
            </td>
           </tr>

           <tr>
            <td>
            As Of Date
            </td>
            <td>
                <asp:TextBox ID="txtAsOfDate" TextMode="Date" runat="server"></asp:TextBox>
            </td>
           </tr>
         
          <tr>
            <td>
                &nbsp;
            </td>
            <td>
                <asp:Button ID="btnGenerate" style="width: 100px;" ClientIDMode="Static" runat="server" Text="Generate" OnClick="btnGenerate_Click"  />
              
            </td>
        </tr>
     </table>

</asp:Content>
