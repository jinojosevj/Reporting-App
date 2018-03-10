<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="VisitorsTati.aspx.cs" Inherits="ReportingTool.VisitorsTati" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>
<asp:Content ID="Content1" ContentPlaceHolderID="phdBodyContent" runat="server">
     <ajaxToolkit:ToolkitScriptManager  ID="ScriptManager1" runat="server"></ajaxToolkit:ToolkitScriptManager> 
    <script type="text/javascript">
        $(document).ready(function () {
            $('#btnReport').click(function () {
                $('#btnReport').hide();

                $('#lblMessage').text("Report Generation Is Going On ...");
                $('#lblMessage').css("color", "Orange");
                $('#lblMessage').show();

            });
        });

      </script>
    
    
     <table width="100%" border="0">
        <tr>
            <td class="text-admin-panel" width="20%">
               TATI Visitors Report
            </td>
            <td>
                <asp:Label ID="lblMessage" runat="server" Text="" ClientIDMode="Static"></asp:Label>
            </td>
        </tr>
    </table>

    <table>
        <tr>
            <td>
            <asp:FileUpload ID="fileuploadExcel" runat="server" />
            </td>
            <td>
            <asp:Button ID="btnImport" runat="server" Text="Import" OnClick="btnImport_Click" />
            </td>
       </tr>
        
        <tr>
             <td>
               Location:
            </td>
             <td>
                 <asp:TextBox ID="txtLocation" runat="server"></asp:TextBox>
            </td>
        </tr>

        <tr>
            <td>
                Posting Date
            </td>
             <td>
                 <asp:TextBox ID="txtDate" runat="server"></asp:TextBox>
                  <ajaxToolkit:CalendarExtender ID="txtDate_CalendarExtender" runat="server" Enabled="True" TargetControlID="txtDate" Format="dd/MM/yyyy" >
                    </ajaxToolkit:CalendarExtender>

            </td>
        </tr>

        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                <asp:Button ID="btnReport" runat="server" Text="Report" OnClick="btnReport_Click" ClientIDMode="Static"/>

                &nbsp;&nbsp; <asp:Button ID="btnDownload"  runat="server" Visible="false" Text="Download Daily" OnClick="btnDownload_Click" Style="width:132px;"/>

               &nbsp;&nbsp; <asp:Button ID="btnDownloadWeekly" runat="server" Text="Download Weekly" Visible="false" Style="width:132px;" OnClick="btnDownloadWeekly_Click"/>

                &nbsp;&nbsp; 
            </td>
        </tr>

        
    </table>

</asp:Content>
