<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="ShipmentReport.aspx.cs" Inherits="ReportingTool.ShipmentReport" %>
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
                Shipment Report TATI
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
                   From Date
                </td>
                <td>
                    <asp:TextBox ID="txtFromDate" runat="server" Text=""></asp:TextBox>
                    <ajaxToolkit:CalendarExtender ID="CalendarExtender4" runat="server" Enabled="True" TargetControlID="txtFromDate" Format="dd/MM/yyyy" >
                    </ajaxToolkit:CalendarExtender>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtFromDate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                </td>
         </tr>

         <tr>
                <td>
                   To Date
                </td>
                <td>
                    <asp:TextBox ID="txtToDate" runat="server" Text=""></asp:TextBox>
                    <ajaxToolkit:CalendarExtender ID="CalendarExtender1" runat="server" Enabled="True" TargetControlID="txtToDate" Format="dd/MM/yyyy" >
                    </ajaxToolkit:CalendarExtender>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtToDate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                </td>
         </tr>
         

        <tr>
            <td>&nbsp;</td>
            <td>
                <asp:Button ID="btnGenerate" runat="server" Text="Generate" style="width: 84px;" ValidationGroup="GenerateReport" ClientIDMode="Static" OnClick="btnGenerate_Click"  />
                &nbsp;<asp:Button ID="btnDownloadShip" runat="server" Text="Download Shipment Report" style="width: 220px;" ValidationGroup="GenerateReport" ClientIDMode="Static" Visible="False" OnClick="btnDownloadShip_Click"  />

            </td>
        </tr>
    </table>

</asp:Content>
