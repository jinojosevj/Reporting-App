<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="DcStock.aspx.cs" Inherits="ReportingTool.DcStock" %>
<asp:Content ID="Content1" ContentPlaceHolderID="phdBodyContent" runat="server">
    <script type="text/javascript">
    function deleteConfirm(pubid) {
        var result = confirm('Do you want to delete ' + pubid + ' ?');
        if (result) {
            return true;
        }
        else {
            return false;
        }
    }
</script> 
    
    <script type="text/javascript">
          $(document).ready(function () {

              $('#mnReport').hide();
              $('#mnTati').hide();
                    
        });
       </script>


    <table style="width: 100%;" border="0">
        <tr>
            <td class="text-admin-panel" width="20%">
                DC Stock
            </td>
            <td>
                 <asp:Label ID="lblMessage" runat="server" ForeColor="Red" ClientIDMode="Static"></asp:Label>&nbsp;
            </td>
        </tr>
    </table>

    <table>
           <tr>
                <td>Type </td>
                <td>
                    <asp:DropDownList ID="ddlType" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlType_SelectedIndexChanged">
                        <asp:ListItem Value="0" Text="Select"></asp:ListItem>
                        <asp:ListItem Value="1" Text="PO Delete/Modify"></asp:ListItem>
                        <asp:ListItem Value="2" Text="SO Delete"></asp:ListItem>
                        <asp:ListItem Value="3" Text="Stock Ledger"></asp:ListItem>
                    </asp:DropDownList>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="ddlType" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                </td>
            </tr>

            <tr runat="server" id="trDocNo" visible="false">
                <td>Document No.</td>
                <td>
                    <asp:DropDownList ID="ddlPONumber"  runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlPONumber_SelectedIndexChanged" >
                        <asp:ListItem Value="0" Text="Select"></asp:ListItem>
                    </asp:DropDownList>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ErrorMessage="This Field is Required" display="Dynamic" ControlToValidate="ddlType" ForeColor="Red" Font-Size="Larger" ValidationGroup="GenerateReport">*</asp:RequiredFieldValidator>
                </td>
            </tr>
        <tr>
            <td colspan="2">
                <asp:Button ID="btnDeletePO" runat="server" visible="false" Text ="Delete PO" style="width: 84px;"  ClientIDMode="Static" OnClick="btnDeletePO_Click"/>
              
                <asp:Button ID="btnDeleteSO" runat="server" visible="false" Text ="Delete SO" style="width: 84px;"  ClientIDMode="Static" OnClick="btnDeleteSO_Click" />

                
            </td>
        </tr>

        <tr runat="server" id="trStockLedger" visible="false">
            <td>
            Stock Ledger &nbsp; <asp:FileUpload ID="fudStockLedger" runat="server" />
                </td>
            <td>
                <asp:Button ID="btnLedger" runat="server" Text ="Ledger Update" style="width: 124px;"  ClientIDMode="Static" OnClick="btnLedger_Click" />
            </td>
        </tr>
        <tr>
            <td colspan="2">
                 <asp:GridView ID="gdvSOData" runat="server">
                 </asp:GridView>
            </td>
        </tr>
     

        <tr>
            <td colspan="2">
            <asp:GridView ID="gridView" DataKeyNames="ID" runat="server"
                AutoGenerateColumns  ="false" ShowFooter="true" HeaderStyle-Font-Bold="true" 
                onrowcancelingedit="gridView_RowCancelingEdit"
                onrowdeleting="gridView_RowDeleting"
                OnRowEditing ="gridView_RowEditing"
                onrowupdating="gridView_RowUpdating"
                onrowcommand="gridView_RowCommand"
                OnRowDataBound="gridView_RowDataBound">
<Columns>
<asp:TemplateField HeaderText="ID">
    <ItemTemplate>
        <asp:Label ID="txtid" runat="server" Text='<%#Eval("ID") %>'/>
    </ItemTemplate>
    <EditItemTemplate>
        <asp:Label ID="lblid"  runat="server" width="40px" Text='<%#Eval("ID") %>'/>
    </EditItemTemplate>
    <FooterTemplate>
        <asp:TextBox ID="inid" ReadOnly="true"  width="40px" runat="server"/>
       
    </FooterTemplate>
</asp:TemplateField>
 <asp:TemplateField HeaderText="PONumber">
      <ItemTemplate>
         <asp:Label ID="lblPONumber" runat="server" Text='<%#Eval("PONumber") %>'/>
     </ItemTemplate>
     <EditItemTemplate>
         <asp:Label ID="txtPONumber" width="55px"  runat="server" Text='<%#Eval("PONumber") %>'/>
     </EditItemTemplate>
     <FooterTemplate>
         <asp:TextBox ID="inPONumber" ReadOnly="true"  width="55px" runat="server"/>
        
     </FooterTemplate>
 </asp:TemplateField>
 <asp:TemplateField HeaderText="LineNo">
     <ItemTemplate>
         <asp:Label ID="lblLineNo" runat="server" Text='<%#Eval("LineNo") %>'/>
     </ItemTemplate>
     <EditItemTemplate>
         <asp:Label ID="txtLineNo" width="30px"  runat="server" Text='<%#Eval("LineNo") %>'/>
     </EditItemTemplate>
    <FooterTemplate>
        <asp:TextBox ID="inLineNo" ReadOnly="true" width="30px"  runat="server"/>
      
    </FooterTemplate>
 </asp:TemplateField>
  <asp:TemplateField HeaderText="LineCode7">
       <ItemTemplate>
         <asp:Label ID="lblLineCode7" runat="server" Text='<%#Eval("LineCode7") %>'/>
     </ItemTemplate>
     <EditItemTemplate>
         <asp:TextBox ID="txtLineCode7" width="50px"   runat="server" Text='<%#Eval("LineCode7") %>'/>
     </EditItemTemplate>
    <FooterTemplate>
        <asp:TextBox ID="inLineCode7" width="60px"  runat="server"/>
        <asp:RequiredFieldValidator ID="vLineCode7" runat="server" ControlToValidate="inLineCode7" Text="?" ValidationGroup="validaiton"/>
    </FooterTemplate>
 </asp:TemplateField>

   <asp:TemplateField HeaderText="PackID">
     <ItemTemplate>
         <asp:Label ID="lblPackID" runat="server" Text='<%#Eval("PackID") %>'/>
     </ItemTemplate>
     <EditItemTemplate>
         <asp:TextBox ID="txtPackID" width="30px"  runat="server" Text='<%#Eval("PackID") %>'/>
     </EditItemTemplate>
    <FooterTemplate>
        <asp:TextBox ID="inPackID" width="40px"   runat="server"/>
        <asp:RequiredFieldValidator ID="vPackID" runat="server" ControlToValidate="inPackID" Text="?" ValidationGroup="validaiton"/>
    </FooterTemplate>
 </asp:TemplateField>
    
    <asp:TemplateField HeaderText="PackBarcode">
     <ItemTemplate>
         <asp:Label ID="lblPackBarcode" runat="server" Text='<%#Eval("PackBarcode") %>'/>
     </ItemTemplate>
    <EditItemTemplate>
         <asp:TextBox ID="txtPackBarcode" width="30px"  runat="server" Text='<%#Eval("PackBarcode") %>'/>
     </EditItemTemplate>
    <FooterTemplate>
        <asp:TextBox ID="inPackBarcode" width="40px"   runat="server"/>
        <asp:RequiredFieldValidator ID="vPackBarcode" runat="server" ControlToValidate="inPackBarcode" Text="?" ValidationGroup="validaiton"/>
    </FooterTemplate>
 </asp:TemplateField>

 <asp:TemplateField HeaderText="PackType">
     <ItemTemplate>
         <asp:Label ID="lblPackType" runat="server" Text='<%#Eval("PackType") %>'/>
     </ItemTemplate>
    <EditItemTemplate>
         <asp:TextBox ID="txtPackType" width="10px"  runat="server" Text='<%#Eval("PackType") %>'/>
     </EditItemTemplate>
    <FooterTemplate>
        <asp:TextBox ID="inPackType" width="10px"   runat="server"/>
        <asp:RequiredFieldValidator ID="vPackType" runat="server" ControlToValidate="inPackType" Text="?" ValidationGroup="validaiton"/>
    </FooterTemplate>
 </asp:TemplateField>

 <asp:TemplateField HeaderText="OrderQty">
     <ItemTemplate>
         <asp:Label ID="lblOrderQty" runat="server" Text='<%#Eval("OrderQty") %>'/>
     </ItemTemplate>
    <EditItemTemplate>
         <asp:TextBox ID="txtOrderQty" width="30px"  runat="server" Text='<%#Eval("OrderQty") %>'/>
     </EditItemTemplate>
    <FooterTemplate>
        <asp:TextBox ID="inOrderQty" width="40px"   runat="server"/>
        <asp:RequiredFieldValidator ID="vinOrderQty" runat="server" ControlToValidate="inOrderQty" Text="?" ValidationGroup="validaiton"/>
    </FooterTemplate>
 </asp:TemplateField>

 <asp:TemplateField HeaderText="UnitPrice">
     <ItemTemplate>
         <asp:Label ID="lblUnitPrice" runat="server" Text='<%#Eval("UnitPrice") %>'/>
     </ItemTemplate>
    <EditItemTemplate>
         <asp:TextBox ID="txtUnitPrice" width="30px"  runat="server" Text='<%#Eval("UnitPrice") %>'/>
     </EditItemTemplate>
    <FooterTemplate>
        <asp:TextBox ID="inUnitPrice" width="40px"   runat="server"/>
        <asp:RequiredFieldValidator ID="vUnitPrice" runat="server" ControlToValidate="inUnitPrice" Text="?" ValidationGroup="validaiton"/>
    </FooterTemplate>
 </asp:TemplateField>


<asp:TemplateField HeaderText="COO">
     <ItemTemplate>
         <asp:Label ID="lblCOO" runat="server" Text='<%#Eval("COO") %>'/>
     </ItemTemplate>
    <EditItemTemplate>
         <asp:TextBox ID="txtCOO" width="30px"  runat="server" Text='<%#Eval("COO") %>'/>
     </EditItemTemplate>
    <FooterTemplate>
        <asp:TextBox ID="inCOO" width="40px"   runat="server"/>
        <asp:RequiredFieldValidator ID="vCOO" runat="server" ControlToValidate="inCOO" Text="?" ValidationGroup="validaiton"/>
    </FooterTemplate>
 </asp:TemplateField>


<asp:TemplateField HeaderText="Department">
     <ItemTemplate>
         <asp:Label ID="lblDepartment" runat="server" Text='<%#Eval("Department") %>'/>
     </ItemTemplate>
    <EditItemTemplate>
         <asp:TextBox ID="txtDepartment" width="30px"  runat="server" Text='<%#Eval("Department") %>'/>
     </EditItemTemplate>
    <FooterTemplate>
        <asp:TextBox ID="inDepartment" width="40px"   runat="server"/>
        <asp:RequiredFieldValidator ID="vDepartment" runat="server" ControlToValidate="inDepartment" Text="?" ValidationGroup="validaiton"/>
    </FooterTemplate>
 </asp:TemplateField>


<asp:TemplateField HeaderText="Nest">
     <ItemTemplate>
         <asp:Label ID="lblNest" runat="server" Text='<%#Eval("Nest") %>'/>
     </ItemTemplate>
    <EditItemTemplate>
         <asp:TextBox ID="txtNest" width="30px"  runat="server" Text='<%#Eval("Nest") %>'/>
     </EditItemTemplate>
    <FooterTemplate>
        <asp:TextBox ID="inNest" width="40px"   runat="server"/>
        <asp:RequiredFieldValidator ID="vNest" runat="server" ControlToValidate="inNest" Text="?" ValidationGroup="validaiton"/>
    </FooterTemplate>
 </asp:TemplateField>


<asp:TemplateField HeaderText="Description">
     <ItemTemplate>
         <asp:Label ID="lblDescription" runat="server" Text='<%#Eval("Description") %>'/>
     </ItemTemplate>
    <EditItemTemplate>
         <asp:TextBox ID="txtDescription" width="30px"  runat="server" Text='<%#Eval("Description") %>'/>
     </EditItemTemplate>
    <FooterTemplate>
        <asp:TextBox ID="inDescription" width="40px"   runat="server"/>
        <asp:RequiredFieldValidator ID="vDescription" runat="server" ControlToValidate="inDescription" Text="?" ValidationGroup="validaiton"/>
    </FooterTemplate>
 </asp:TemplateField>


<asp:TemplateField HeaderText="Season">
     <ItemTemplate>
         <asp:Label ID="lblSeason" runat="server" Text='<%#Eval("Season") %>'/>
     </ItemTemplate>
    <EditItemTemplate>
         <asp:TextBox ID="txtSeason" width="30px"  runat="server" Text='<%#Eval("Season") %>'/>
     </EditItemTemplate>
    <FooterTemplate>
        <asp:TextBox ID="inSeason" width="40px"   runat="server"/>
        <asp:RequiredFieldValidator ID="vSeason" runat="server" ControlToValidate="inSeason" Text="?" ValidationGroup="validaiton"/>
    </FooterTemplate>
 </asp:TemplateField>

<asp:TemplateField HeaderText="Outer">
     <ItemTemplate>
         <asp:Label ID="lblOuter" runat="server" Text='<%#Eval("Outer") %>'/>
     </ItemTemplate>
    <EditItemTemplate>
         <asp:TextBox ID="txtOuter" width="30px"  runat="server" Text='<%#Eval("Outer") %>'/>
     </EditItemTemplate>
    <FooterTemplate>
        <asp:TextBox ID="inOuter" width="40px"   runat="server"/>
        <asp:RequiredFieldValidator ID="vOuter" runat="server" ControlToValidate="inOuter" Text="?" ValidationGroup="validaiton"/>
    </FooterTemplate>
 </asp:TemplateField>

<asp:TemplateField HeaderText="Invoiced">
     <ItemTemplate>
         <asp:Label ID="lblInvoiced" runat="server" Text='<%#Eval("Invoiced") %>'/>
     </ItemTemplate>
    <EditItemTemplate>
         <asp:TextBox ID="txtInvoiced" width="30px"  runat="server" Text='<%#Eval("Invoiced") %>'/>
     </EditItemTemplate>
    <FooterTemplate>
        <asp:TextBox ID="inInvoiced" width="40px"   runat="server"/>
        <asp:RequiredFieldValidator ID="vInvoiced" runat="server" ControlToValidate="inInvoiced" Text="?" ValidationGroup="validaiton"/>
    </FooterTemplate>
 </asp:TemplateField>

<asp:TemplateField HeaderText="PackLevel">
     <ItemTemplate>
         <asp:Label ID="lblPackLevel" runat="server" Text='<%#Eval("PackLevel") %>'/>
     </ItemTemplate>
    <EditItemTemplate>
         <asp:TextBox ID="txtPackLevel" width="30px"  runat="server" Text='<%#Eval("PackLevel") %>'/>
     </EditItemTemplate>
    <FooterTemplate>
        <asp:TextBox ID="inPackLevel" width="40px"   runat="server"/>
        <asp:RequiredFieldValidator ID="vPackLevel" runat="server" ControlToValidate="inPackLevel" Text="?" ValidationGroup="validaiton"/>
    </FooterTemplate>
 </asp:TemplateField>


 <asp:TemplateField>
    <EditItemTemplate>
        <asp:Button ID="ButtonUpdate" runat="server" CommandName="Update"  Text="Update"  />
        <asp:Button ID="ButtonCancel" runat="server" CommandName="Cancel"  Text="Cancel" />
    </EditItemTemplate>
    <ItemTemplate>
        <asp:Button ID="ButtonEdit" runat="server" CommandName="Edit"  Text="Edit"  />
       
    </ItemTemplate>
    <FooterTemplate>
        <asp:Button ID="ButtonAdd" runat="server" CommandName="AddNew"  Text="Add New Row" ValidationGroup="validaiton" />
    </FooterTemplate>
 </asp:TemplateField>
 </Columns>
</asp:GridView>
       
 </td>

        </tr>
    </table>

</asp:Content>
