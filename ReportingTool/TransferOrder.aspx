<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="TransferOrder.aspx.cs" Inherits="ReportingTool.TransferOrder" %>
<asp:Content ID="Content1" ContentPlaceHolderID="phdBodyContent" runat="server">

      <script type="text/javascript">
         $(window).ready(function () {
             $('#loading').hide();
         });

         $(document).ready(function () {
             $('#btnImport').click(function () {
                 $('#btnImport').hide();
                 $('#btnDownload').hide();
                 
                 $('#lblMessage').text("Report Generation Is Going On ...");
                 $('#lblMessage').css("color", "Orange");
                 $('#lblMessage').show();

                 $('#loading').show();
             });

             $('#btnImportReceiver').click(function () {
                 $('#btnImportReceiver').hide();
                 $('#btnDownload').hide();

                 $('#lblMessage').text("Report Generation Is Going On ...");
                 $('#lblMessage').css("color", "Orange");
                 $('#lblMessage').show();

                 $('#loading').show();
             });

             $('#btnGenerate').click(function () {
                 $('#btnGenerate').hide();
                 $('#btnDownload').hide();

                 $('#lblMessage').text("Report Generation Is Going On ...");
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
               Transfer Order
            </td>
            <td>
                 <asp:Label ID="lblMessage" runat="server" ForeColor="Red" ClientIDMode="Static"></asp:Label>&nbsp;
            </td>
        </tr>
    </table>

     <asp:Panel ID="Panel1" runat="server" GroupingText="Step1:Transfer Report">
    <table>
          
          <tr>
            <td>
                Doc No. :
            </td>
            <td>
                 <asp:TextBox ID="txtDocumentNo" runat="server"></asp:TextBox>
            </td>

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
                    </asp:DropDownList>
            </td>
            <td>
                &nbsp;
            </td>
        </tr>

        <tr>
                <td style="height: 32px">
                  Sender Location :
                </td>
                <td style="height: 32px">
                    <asp:DropDownList ID="ddlLocation" runat="server" >
                        <asp:ListItem Value="Select" Text="Select"></asp:ListItem>
                    </asp:DropDownList>
                </td>
                 <td>
                    &nbsp;
                 </td>
            </tr>

       <tr>
            <td>
               Sender  File :
            </td>
            <td>
                  <asp:FileUpload ID="fileuploadSender"  runat="server" /> 
            </td>

             <td>
                  <asp:Button ID="btnImport" ClientIDMode="Static" style="width: 150px;" runat="server" Text="Import - Sender" OnClick="btnImport_Click"  />
             </td>
        </tr>
        <tr>
            <td>
               Receiver File :
            </td>
            <td>
                  <asp:FileUpload ID="fileuploadReceiver"  runat="server" /> 
            </td>
             <td>
                  <asp:Button ID="btnImportReceiver" ClientIDMode="Static" style="width: 150px;" runat="server" Text="Import - Receiver" OnClick="btnImportReceiver_Click" />
             </td>
        </tr>

       

        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                <asp:Button ID="btnGenerate" ClientIDMode="Static" style="width: 100px;" runat="server" Text="Generate" OnClick="btnGenerate_Click"  />
                &nbsp;&nbsp;
                <asp:Button ID="btnDownload" ClientIDMode="Static" style="width: 100px;" runat="server" Text="Download" Visible="false" OnClick="btnDownload_Click" />
            </td>
        </tr>
     </table>

      </asp:Panel>
    <br />
    <asp:Panel ID="Panel2" runat="server" GroupingText="Step2: Post To Nav">
                <table>
                    <tr>
                        <td>Import Adjustment:</td>
                        <td><asp:FileUpload ID="fileuploadAdjustment"  runat="server" /> </td>
                         <td>
                              <asp:Button ID="btnImportAdjustment" ClientIDMode="Static" style="width: 150px;" runat="server" Text="Import - Adjustment" OnClick="btnImportAdjustment_Click" />
                        </td>
                    </tr>
                     <tr>
                        <td>&nbsp;</td>
                        
                         <td><asp:Button ID="btnPost" ClientIDMode="Static" style="width: 100px;" runat="server" Text="Post"  OnClick="btnPost_Click" /></td>
                    </tr>
                </table>
     </asp:Panel>

</asp:Content>
