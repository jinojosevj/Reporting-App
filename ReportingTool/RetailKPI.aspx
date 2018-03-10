﻿<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="RetailKPI.aspx.cs" Inherits="ReportingTool.RetailKPI" %>
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
                Retail KPI Report
            </td>
            <td>
                 <asp:Label ID="lblMessage" runat="server" ForeColor="Red" ClientIDMode="Static"></asp:Label>&nbsp;
            </td>
        </tr>
    </table>

    <table>
         <tr>
            <td>
           Sales Plan&nbsp; <asp:FileUpload ID="fileuploadExcel" runat="server" />
            </td>
            <td>
            <asp:Button ID="btnImport" style="width: 100px;" runat="server" Text="Import Plan" OnClick="btnImport_Click" />
            </td>
        </tr>

        <tr>
            <td>
           Linear Count&nbsp; <asp:FileUpload ID="fudLinearCount" runat="server" />
            </td>
            <td>
            <asp:Button ID="btnImportLinear" style="width: 105px;" runat="server" Text="Import Linear" OnClick="btnImportLinear_Click"/>
            </td>
        </tr>

         <tr>
                <td>
                    Posting Date
                </td>
                <td>
                    <asp:TextBox ID="txtPostingDate" runat="server"></asp:TextBox>
                    <ajaxToolkit:CalendarExtender ID="CalendarExtender5" runat="server" Enabled="True" TargetControlID="txtPostingDate" Format="dd/MM/yyyy" >
                    </ajaxToolkit:CalendarExtender>
                     <asp:RequiredFieldValidator ID="RequiredFieldValidator16"  runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtPostingDate" ForeColor="Red" Font-Size="Larger" ValidationGroup="AddSales">*</asp:RequiredFieldValidator>
                </td>
                
            </tr>
        <tr>
            <td>
               Sales Addition / Deduction 
            </td>
            <td>
                 <asp:TextBox ID="txtSalesAmount" runat="server"></asp:TextBox>
                <asp:RequiredFieldValidator ID="RequiredFieldValidator15"  runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtSalesAmount" ForeColor="Red" Font-Size="Larger" ValidationGroup="AddSales">*</asp:RequiredFieldValidator>
            </td>
        </tr>
        <tr>
            <td>
                Location
            </td>
            <td>

                <asp:DropDownList ID="ddlLocation" runat="server">
                        <asp:ListItem Value="0400" Text="0400"></asp:ListItem>
                        <asp:ListItem Value="0401" Text="0401"></asp:ListItem>
                        <asp:ListItem Value="0402" Text="0402"></asp:ListItem>
                        <asp:ListItem Value="0403" Text="0403"></asp:ListItem>

                        <asp:ListItem Value="0404" Text="0404"></asp:ListItem>
                        <asp:ListItem Value="0405" Text="0405"></asp:ListItem>
                        <asp:ListItem Value="0406" Text="0406"></asp:ListItem>
                        <asp:ListItem Value="0407" Text="0407"></asp:ListItem>

                        <asp:ListItem Value="0409" Text="0409"></asp:ListItem>
                        <asp:ListItem Value="0410" Text="0410"></asp:ListItem>
                        <asp:ListItem Value="0411" Text="0411"></asp:ListItem>
                        <asp:ListItem Value="0412" Text="0412"></asp:ListItem>

                        <asp:ListItem Value="0414" Text="0414"></asp:ListItem>
                        <asp:ListItem Value="0415" Text="0415"></asp:ListItem>
                        <asp:ListItem Value="0416" Text="0416"></asp:ListItem>
                        <asp:ListItem Value="0417" Text="0417"></asp:ListItem>

                        <asp:ListItem Value="0418" Text="0418"></asp:ListItem>
                        <asp:ListItem Value="0419" Text="0419"></asp:ListItem>
                        <asp:ListItem Value="0421" Text="0421"></asp:ListItem>
                        <asp:ListItem Value="0424" Text="0424"></asp:ListItem>

                        <asp:ListItem Value="0425" Text="0425"></asp:ListItem>
                        <asp:ListItem Value="0426" Text="0426"></asp:ListItem>
                    </asp:DropDownList>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator14"  runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="ddlLocation" ForeColor="Red" Font-Size="Larger" ValidationGroup="AddSales">*</asp:RequiredFieldValidator>
            </td>
        </tr>
         <tr>
            <td>
                Description
            </td>
            <td>
                <asp:TextBox ID="txtDescription"   runat="server"></asp:TextBox>
                 <asp:RequiredFieldValidator ID="RequiredFieldValidator17"  runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtDescription" ForeColor="Red" Font-Size="Larger" ValidationGroup="AddSales">*</asp:RequiredFieldValidator>
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                 <asp:Button ID="btnAddSales" style="width: 70px;" runat="server" Text="Add" ValidationGroup="AddSales" OnClick="btnAddSales_Click"/>
            </td>
        
        </tr>

         <tr>
                <td>
                    Retail KPI Table
                </td>
                <td>
                     <asp:RadioButtonList ID="rdlRetailKpi" runat="server" RepeatDirection="Horizontal">
                         <asp:ListItem Value="1" Selected="True">Refresh</asp:ListItem>
                         <asp:ListItem Value="0">No Refresh</asp:ListItem>
                     </asp:RadioButtonList>
                </td>
            </tr>


          <tr>
                <td>Report Date</td>
                <td>
                    
                    <asp:TextBox ID="txtReportDate" runat="server"></asp:TextBox>
                    <ajaxToolkit:CalendarExtender ID="CalendarExtender4" runat="server" Enabled="True" TargetControlID="txtReportDate" Format="dd/MM/yyyy" >
                    </ajaxToolkit:CalendarExtender>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator12" display="Dynamic" runat="server" ErrorMessage="This Field is Required" ControlToValidate="txtReportDate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
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
                    <asp:TextBox ID="txtKsaRate" runat="server" Text="0.967"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator13" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtKsaRate" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator4" ControlToValidate="txtKsaRate" runat="server"
                        ErrorMessage="Number Only" Display="Dynamic" ValidationExpression="[0-9]*\.?[0-9]*"></asp:RegularExpressionValidator>
                
                </td>
             </tr>


           <tr>
                <td>
                   Year Start Date</td>
                <td>
                      <asp:TextBox ID="txtYearStart" runat="server"></asp:TextBox>
                        <ajaxToolkit:CalendarExtender ID="CalendarExtender3" runat="server" Enabled="True" TargetControlID="txtYearStart" Format="dd/MM/yyyy" >
                        </ajaxToolkit:CalendarExtender>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator11" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtYearStart" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                 </td>
            </tr>


          <tr>
                <td>
                   Month Start Date</td>
                <td>
                      <asp:TextBox ID="txtMonthStart" runat="server"></asp:TextBox>
                        <ajaxToolkit:CalendarExtender ID="CalendarExtender1" runat="server" Enabled="True" TargetControlID="txtMonthStart" Format="dd/MM/yyyy" >
                        </ajaxToolkit:CalendarExtender>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator9" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtMonthStart" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                 </td>
            </tr>
          
         <tr>
                <td>
                   Month End Date</td>
                <td>
                      <asp:TextBox ID="txtMonthEnd" runat="server"></asp:TextBox>
                        <ajaxToolkit:CalendarExtender ID="CalendarExtender2" runat="server" Enabled="True" TargetControlID="txtMonthEnd" Format="dd/MM/yyyy" >
                        </ajaxToolkit:CalendarExtender>
                      <asp:RequiredFieldValidator ID="RequiredFieldValidator10" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtMonthEnd" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                 </td>
            </tr>
            
            <tr>
                <td>
                   Current Week No
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
                        <asp:ListItem Value="2018" Text="2018-2019"></asp:ListItem>
                        <asp:ListItem Value="2019" Text="2019-2020"></asp:ListItem>
                    </asp:DropDownList>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator8" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="txtWeekNo" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                </td>
            </tr>

        <tr>
            <td>&nbsp;</td>
            <td>
                <asp:Button ID="btnGenerate" runat="server" Text="Generate" style="width: 84px;" ValidationGroup="GenerateReport" ClientIDMode="Static" OnClick="btnGenerate_Click"  />
                &nbsp;<asp:Button ID="btnDownload" runat="server" Text="Retail KPI" style="width: 94px;" ValidationGroup="GenerateReport" ClientIDMode="Static" Visible="False" OnClick="btnDownload_Click"/>
                &nbsp;<asp:Button ID="btnDownloadYear" runat="server" Text="Retail KPI Year" style="width: 115px;" ValidationGroup="GenerateReport" ClientIDMode="Static" Visible="False" OnClick="btnDownloadYear_Click" />
                &nbsp;<asp:Button ID="btnDownloadDsr" runat="server" Text="DSR" style="width: 55px;" ValidationGroup="GenerateReport" ClientIDMode="Static" Visible="False" OnClick="btnDownloadDsr_Click" />
            </td>
        </tr>
    </table>
</asp:Content>