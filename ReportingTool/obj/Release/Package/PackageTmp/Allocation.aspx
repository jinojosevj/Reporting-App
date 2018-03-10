<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Allocation.aspx.cs" Inherits="ReportingTool.Allocation" %>
<asp:Content ID="Content1" ContentPlaceHolderID="phdBodyContent" runat="server">
     <script type="text/javascript">
         $(window).ready(function () {
             $('#loading').hide();
         });

         $(document).ready(function () {
             $('#btnGenerate').click(function () {
                 $('#btnGenerate').hide();
                 $('#btnDownloadDC').hide();
                 $('#btnDownloadDCS').hide();
                
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
                 DC Allocation Process
            </td>
            <td>
                 <asp:Label ID="lblMessage" runat="server" ForeColor="Red" ClientIDMode="Static"></asp:Label>&nbsp;
            </td>
        </tr>
    </table>
      

       <table>
          
          <%-- <tr>
                <td>
                  Mode :
                </td>
                <td>
                      <asp:DropDownList ID="ddlMode" runat="server">
                        <asp:ListItem Value="1" Text="Single"></asp:ListItem>
                        <asp:ListItem Value="2" Text="Pack"></asp:ListItem>
                        
                    </asp:DropDownList>
                </td>
            </tr>--%>

           <tr>
                <td>
                 Sell Through(%) :
                </td>
                <td>
                      <asp:TextBox ID="txtSellThrough" TextMode="Number"  runat="server"></asp:TextBox>
                </td>
            </tr>

           <tr>
                <td>
                  Store Code :
                </td>
                <td>
                      <asp:TextBox ID="txtStoreCode"  runat="server"></asp:TextBox>
                </td>
            </tr>

           
          <tr>
            <td>
                &nbsp;
            </td>
            <td>
                <asp:Button ID="btnGenerate" style="width: 100px;" runat="server" ClientIDMode="Static" Text="Generate" OnClick="btnGenerate_Click"  />
              
                 &nbsp;&nbsp;
                <asp:Button ID="btnDownloadDC" style="width: 150px;" runat="server" ClientIDMode="Static" Text="Dwd DCStock-Pack" Visible="false" OnClick="btnDownloadDC_Click" />
                 &nbsp;&nbsp;
                <asp:Button ID="btnDownloadDCS" style="width: 150px;" runat="server" ClientIDMode="Static" Text="Dwd DCStock-Single" Visible="false" OnClick="btnDownloadDCS_Click" />

            </td>
        </tr>
     </table>
   
</asp:Content>
