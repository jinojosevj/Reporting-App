﻿<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="StockStatusTati.aspx.cs" Inherits="ReportingTool.StockStatusTati" %>
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
               TATI Stock Status Report
            </td>
            <td>
                
                 <asp:Label ID="lblMessage" runat="server" ForeColor="Red" ClientIDMode="Static"></asp:Label>&nbsp;
               
            </td>
        </tr>
    </table>
    
    
        <table style="width: 100%;">
            <tr>
                <td>
                    Stock Status Table
                </td>
                <td>
                     <asp:RadioButtonList ID="rdlStockStatus" runat="server" RepeatDirection="Horizontal">
                         <asp:ListItem Value="1" Selected="True">Refresh</asp:ListItem>
                         <asp:ListItem Value="0">No Refresh</asp:ListItem>
                     </asp:RadioButtonList>
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
             
            <tr runat="server" id="tdLocation" >
                <td>Location</td>
                <td>
                    <asp:TextBox ID="txtLocation" runat="server"></asp:TextBox></td>
            </tr>

            <tr>
                <td>UAE Exchange Rate</td>
                <td>
                    <asp:TextBox ID="txtUaeRate" runat="server" Text="0.99" ></asp:TextBox>
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
                    <asp:TextBox ID="txtKSARate" runat="server" Text="0.967"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator8" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtBahrainRate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator4" ControlToValidate="txtKSARate" runat="server"
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
                <td>
                   Year
                </td>
                <td>
                    <asp:DropDownList ID="ddlYear" runat="server">
                        <asp:ListItem Value="2014" Text="2014-2015"></asp:ListItem>
                        <asp:ListItem Value="2015" Text="2015-2016"></asp:ListItem>
                        <asp:ListItem Value="2016" Text="2016-2017"></asp:ListItem>
                        <asp:ListItem Value="2017" Text="2017-2018"></asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
           


            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                    <asp:Button ID="btnGenerate" runat="server" Text="Generate" style="width: 84px;" ValidationGroup="GenerateReport" OnClick="btnGenerate_Click" ClientIDMode="Static"/>
                    &nbsp;<asp:Button ID="btnDownload" runat="server" Text="Download-SSR" style="width: 110px;"  OnClick="btnDownload_Click" ClientIDMode="Static" Visible="false"/>

                    &nbsp;<asp:Button ID="btnDownloadLCP" runat="server" Text="Download-LCP" style="width: 110px;"  ClientIDMode="Static" Visible="false" OnClick="btnDownloadLCP_Click"/>

                </td>
                
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                    
                      <asp:Button ID="btnSSRMY" runat="server" Text="Download-SSR-MY" style="width: 170px;"   ClientIDMode="Static" Visible="False" OnClick="btnSSRMY_Click" />
                    
                     &nbsp;<asp:Button ID="btnLCPMY" runat="server" Text="Download-LCP-MY" style="width: 160px;"   ClientIDMode="Static" Visible="false" OnClick="btnLCPMY_Click" />
                  
                   
                </td>
                
            </tr>

            <tr>
                <td>
                    &nbsp;
                </td>

                <td>
                     
                </td>

            </tr>
        </table>

</asp:Content>
