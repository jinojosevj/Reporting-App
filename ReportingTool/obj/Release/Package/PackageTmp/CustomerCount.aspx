<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="CustomerCount.aspx.cs" Inherits="ReportingTool.CustomerCount" %>
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
                Customer Count Report
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
            <td>From Date:</td>
            <td>
                    <asp:TextBox ID="txtFromDate" runat="server"></asp:TextBox>
                    <ajaxtoolkit:calendarextender ID="txtAsofDate_CalendarExtender" runat="server" Enabled="True" TargetControlID="txtFromDate" Format="dd/MM/yyyy" >
                    </ajaxtoolkit:calendarextender>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" display="Dynamic" runat="server" ErrorMessage="This Field is Required" ControlToValidate="txtFromDate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
            </td>
        </tr>

        <tr>
            <td>To Date:</td>
            <td>
                    <asp:TextBox ID="txtToDate" runat="server"></asp:TextBox>
                    <ajaxtoolkit:calendarextender ID="Calendarextender1" runat="server" Enabled="True" TargetControlID="txtToDate" Format="dd/MM/yyyy" >
                    </ajaxtoolkit:calendarextender>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" display="Dynamic" runat="server" ErrorMessage="This Field is Required" ControlToValidate="txtToDate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
            </td>
        </tr>

        <tr>
            <td>&nbsp;</td>
            <td>
                    <asp:Button ID="btnGenerate" runat="server" Text="Generate" style="width: 84px;" ValidationGroup="GenerateReport" OnClick="btnGenerate_Click" ClientIDMode="Static"/>
                    
                   &nbsp; <asp:Button ID="btnDownload" runat="server" Text="Customer Count Report" style="width: 175px;"  Visible="false" ClientIDMode="Static" OnClick="btnDownload_Click" />

            </td>
        </tr>
    </table>

</asp:Content>
