<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="MISReports.aspx.cs" Inherits="ReportingTool.MISReports" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>
<asp:Content ID="Content1" ContentPlaceHolderID="phdBodyContent" runat="server">
     <ajaxToolkit:ToolkitScriptManager  ID="ScriptManager1" runat="server"></ajaxToolkit:ToolkitScriptManager> 
     <script type="text/javascript">
         $(window).ready(function () {
             $('#loading').hide();
         });

         $(document).ready(function () {
             var startTime, endTime;

             
             $('#btnGenerate').click(function () {
                
                 $('#btnGenerate').hide();
                 $('#btnDownload').hide();
                 
                 $('#lblMessage').text("Report Generation Is Going On ..." );
                 $('#lblMessage').css("color", "Orange");
                 $('#lblMessage').show();

                 $('#loading').show();
             });
         });


       </script>
     <script type="text/javascript">
          $(document).ready(function () {
              $('#mnReport').hide();
              $('#mnTati').hide();
        });
       </script>


     <div  id="loading" ></div>
    <table style="width: 100%;" border="0">
        <tr>
            <td class="text-admin-panel" width="20%">
                MIS Reports
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
                        <asp:ListItem Value="1" Text="GBP Price"></asp:ListItem>
                        <asp:ListItem Value="2" Text="Sales"></asp:ListItem>
                        <asp:ListItem Value="3" Text="Cost"></asp:ListItem>
                        <asp:ListItem Value="4" Text="Family Code"></asp:ListItem>

                        <asp:ListItem Value="5" Text="Inventory Summary"></asp:ListItem>
                        <asp:ListItem Value="6" Text="Stock Cover"></asp:ListItem>
                        <asp:ListItem Value="7" Text="Store Stock"></asp:ListItem>
                        <asp:ListItem Value="8" Text="Sales By Product Group"></asp:ListItem>
                        <asp:ListItem Value="9" Text="Sales By Item Family"></asp:ListItem>

                        <asp:ListItem Value="10" Text="WSSI"></asp:ListItem>
                        <asp:ListItem Value="11" Text="Item Sales Price"></asp:ListItem>
                        <asp:ListItem Value="12" Text="Inventory For Promotion"></asp:ListItem>
                        <asp:ListItem Value="13" Text="Update Unit Price"></asp:ListItem>

                        <asp:ListItem Value="14" Text="Update Tables"></asp:ListItem>
                        <asp:ListItem Value="15" Text="Update Sales Line"></asp:ListItem>
                        <asp:ListItem Value="16" Text="Inventory Adjustment"></asp:ListItem>
                        <asp:ListItem Value="17" Text="Markdown Report"></asp:ListItem>

                        <asp:ListItem Value="18" Text="Sellthrough"></asp:ListItem>
                        <asp:ListItem Value="19" Text="Unit Cost For Stock Count"></asp:ListItem>
                       
                    </asp:DropDownList>
            </td>
        </tr>

         <tr>
            <td>
                Country:
            </td>
            <td>
                   <asp:DropDownList ID="ddlCountry" runat="server" OnSelectedIndexChanged="ddlCountry_SelectedIndexChanged" AutoPostBack="True">
                        <asp:ListItem Value="Select" Text="Select"></asp:ListItem>
                        <asp:ListItem Value="UAE" Text="UAE"></asp:ListItem>
                        <asp:ListItem Value="JOR" Text="JORDAN"></asp:ListItem>
                        <asp:ListItem Value="OMAN" Text="OMAN"></asp:ListItem>
                        <asp:ListItem Value="BAH" Text="BAHRAIN"></asp:ListItem>
                        <asp:ListItem Value="QAT" Text="QATAR"></asp:ListItem>
                        <asp:ListItem Value="KSA" Text="KSA"></asp:ListItem>
                        <asp:ListItem Value="DC" Text="DC"></asp:ListItem>
                        <asp:ListItem Value="MME" Text="MME"></asp:ListItem>
                    </asp:DropDownList>
            </td>
        </tr>
          <tr>
                <td style="height: 32px">
                   Location :
                </td>
                <td style="height: 32px">
                    <asp:DropDownList ID="ddlLocation" runat="server" >
                        <asp:ListItem Value="Select" Text="Select"></asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
          
           <tr>
                <td>
                   Division Code :
                </td>
                <td>
                      <asp:DropDownList ID="ddlDivision" runat="server" >
                        <asp:ListItem Value="Select" Text="Select"></asp:ListItem>
                        <asp:ListItem Value="C" Text="CHILDRENSWEAR"></asp:ListItem>
                        <asp:ListItem Value="F" Text="FOOTWEAR AND ACCESSORIES"></asp:ListItem>
                        <asp:ListItem Value="H" Text="HOMEWARE"></asp:ListItem>
                        <asp:ListItem Value="L" Text="LADIESWEAR"></asp:ListItem>
                        <asp:ListItem Value="M" Text="MENSWEAR"></asp:ListItem>
                        <asp:ListItem Value="P" Text="PROMOTIONAL"></asp:ListItem>

                        <asp:ListItem Value="R" Text="SPORTS"></asp:ListItem>
                        <asp:ListItem Value="S" Text="OWN BRAND SPORTS"></asp:ListItem>
                        <asp:ListItem Value="Z" Text="OTHERS"></asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
        
         
           <tr>
                <td>
                   LineCode7/12 :
                </td>
                <td>
                      <asp:TextBox ID="txtLinecode7" TextMode="MultiLine" Rows="5"  Columns="30" runat="server"></asp:TextBox>
                </td>
            </tr>

           <tr>
            <td>
            From Date
            </td>
            <td>
                <asp:TextBox ID="txtFromDate" TextMode="Date" runat="server"></asp:TextBox>
            </td>
        </tr>

          <tr>
            <td>
            To Date
            </td>
            <td>
                <asp:TextBox ID="txtToDate"  TextMode="Date" runat="server"></asp:TextBox>
            </td>
        </tr>
         
          <tr>
            <td>
                Unit Price
            </td>
            <td>
                  <asp:FileUpload ID="fileuploadExcel"  runat="server" /> 
            </td>
        </tr>

        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                <asp:Button ID="btnGenerate" ClientIDMode="Static" style="width: 100px;" runat="server" Text="Generate" OnClick="btnGenerate_Click" />
                &nbsp;&nbsp;
                <asp:Button ID="btnDownload" ClientIDMode="Static" style="width: 100px;" runat="server" Text="Download" Visible="false" OnClick="btnDownload_Click" />
            </td>
        </tr>
     </table>
</asp:Content>
