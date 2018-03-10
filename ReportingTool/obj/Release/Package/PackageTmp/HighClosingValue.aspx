<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="HighClosingValue.aspx.cs" Inherits="ReportingTool.HighClosingValue" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>
<asp:Content ID="Content1" ContentPlaceHolderID="phdBodyContent" runat="server">
    <ajaxtoolkit:toolkitscriptmanager  ID="ScriptManager1" runat="server"></ajaxtoolkit:toolkitscriptmanager>
    
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
    
    <table width="100%" border="0">
        <tr>
            <td class="text-admin-panel" width="20%">
                High Closing Value Report
            </td>
            <td>
                
                 <asp:Label ID="lblMessage" runat="server" ForeColor="Red" ClientIDMode="Static"></asp:Label>&nbsp;
               
            </td>
        </tr>
    </table>
    
    <table  style="width:100%; border:0">
        <tr>
            <td>Location Code :</td>
            <td><asp:TextBox ID="txtLocation" runat="server"></asp:TextBox></td>
        </tr>
        <tr>
            <td>Report Type:</td>
            <td>
                    <asp:DropDownList ID="ddlType" runat="server">
                        <asp:ListItem Value="CV" Text="Cost Value"></asp:ListItem>
                        <asp:ListItem Value="RV" Text="Retail Value"></asp:ListItem>
                        <asp:ListItem Value="UC" Text="Unit Cost"></asp:ListItem>
                    </asp:DropDownList>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" display="Dynamic" runat="server" ErrorMessage="This Field is Required" ControlToValidate="ddlType" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
            </td>
        </tr>

        <tr>
            <td>As Of Date:</td>
            <td>
                    <asp:TextBox ID="txtAsOfDate" runat="server"></asp:TextBox>
                    <ajaxtoolkit:calendarextender ID="Calendarextender1" runat="server" Enabled="True" TargetControlID="txtAsOfDate" Format="dd/MM/yyyy" >
                    </ajaxtoolkit:calendarextender>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" display="Dynamic" runat="server" ErrorMessage="This Field is Required" ControlToValidate="txtAsOfDate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
            </td>
        </tr>

        <tr>
            <td>&nbsp;</td>
            <td>
                    <asp:Button ID="btnGenerate" runat="server" Text="Generate" style="width: 84px;" ValidationGroup="GenerateReport" ClientIDMode="Static" OnClick="btnGenerate_Click"/>
                    
                   &nbsp; <asp:Button ID="btnDownload" runat="server" Text="Closing Value Report" style="width: 175px;"  Visible="false" ClientIDMode="Static" OnClick="btnDownload_Click" />

            </td>
        </tr>
    </table>


</asp:Content>
