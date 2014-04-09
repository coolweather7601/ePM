<%@ Page Language="C#" MasterPageFile="~/View/Master/MasterPage.master" AutoEventWireup="true" CodeFile="Item.aspx.cs" Inherits="EPM_Web.Alan.Item" Title="ePM" %>
<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>
    
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder2" Runat="Server">

<asp:UpdatePanel id="up" runat="server">
    <Triggers>
        <asp:PostBackTrigger ControlID="btnImportCsv" />
    </Triggers>
    <ContentTemplate>
    
    <h2><%=sheet_Category_describe %>_Item</h2>
     

    <table>
        <tr>
            <td>Describe：</td>
            <td><asp:TextBox ID="txtDescribe" runat="server"></asp:TextBox></td>
            <td>Method：</td>
            <td><asp:TextBox ID="txtMethod" runat="server"></asp:TextBox></td>
            <td>Spec：</td>
            <td><asp:TextBox ID="txtSpec" runat="server"></asp:TextBox></td>
        </tr>
        <tr>
            <td>isNumeric：</td>
            <td><asp:DropDownList id="ddlIsNumeric" runat="server"></asp:DropDownList></td>
            <td>isUrgent：</td>
            <td><asp:DropDownList id="ddlIsUrgent" runat="server"></asp:DropDownList></td>
            <td colspan="2" style="text-align:right;"><asp:Button ID="btnSearch" runat="server" Text="Search" OnClick="btnSearch_Click" /></td>
        </tr> 
        <tr>
            <td>
                <asp:HyperLink ID="HyperLink1" NavigateUrl="~/download/items_import.csv" runat="server">Import Template</asp:HyperLink>                
            </td>
            <td colspan="5">
                <asp:FileUpload ID="FileUpload1" runat="server" />
                <asp:Button ID="btnImportCsv" runat="server" Text="Import" onclick="btnImportCsv_Click" />
            </td>
        </tr>
        <tr>    
            <td colspan="6">
                <asp:Button ID="btnNew" runat="server" Text="New Data" OnClick="btnNew_Click" />
            </td>
        </tr>       
    </table> 
      
    <asp:GridView ID="GridView1" runat="server" DataKeyNames="ItemID" AutoGenerateColumns="False" 
                      CellPadding="3" ForeColor="Black" GridLines="Vertical" AllowSorting="True" AllowPaging="True"
            OnSorting="GridView1_Sorting" 
            OnRowCancelingEdit="GridView1_RowCancelingEdit" 
            OnRowEditing="GridView1_RowEditing" 
            OnRowUpdating="GridView1_RowUpdating" 
            OnRowDataBound="GridView1_RowDataBound"  
            OnPageIndexChanging="GridView1_PageIndexChanging" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" OnRowCommand="GridView1_RowCommand">
            <Columns>                
                <asp:TemplateField HeaderText="No.">
                        <ItemTemplate>
                            <%# (Container.DataItemIndex +1).ToString()%>
                        </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="DataType_desc" SortExpression="datatype_desc">
                    <EditItemTemplate>
                        <asp:DropDownList id="gv_ddlDatatype" runat="server" AutoPostBack="True" OnSelectedIndexChanged="gv_ddlDatatype_SelectedIndexChanged"></asp:DropDownList>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label8" runat="server" Text='<%# Bind("datatype_desc") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="IsNumeric" SortExpression="isnumeric">
                    <EditItemTemplate>
                        <asp:RadioButton id="gv_rdoIsNumeric_Y" GroupName="B" Text="Yes" runat="server" AutoPostBack="True" OnCheckedChanged="gv_rdoIsNumeric_Y_CheckedChanged"></asp:RadioButton>
                        <br />
                        <asp:RadioButton id="gv_rdoIsNumeric_N" GroupName="B" Text="No" runat="server" AutoPostBack="True" OnCheckedChanged="gv_rdoIsNumeric_N_CheckedChanged"></asp:RadioButton>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label7" runat="server" Text='<%# Bind("isnumeric") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="IsUrgent" SortExpression="isurgent">
                    <EditItemTemplate>
                        <asp:RadioButton id="gv_rdoIsUrgent_Y" GroupName="A" Text="Yes" runat="server"></asp:RadioButton>
                        <br />
                        <asp:RadioButton id="gv_rdoIsUrgent_N" GroupName="A" Text="No" runat="server"></asp:RadioButton>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label6" runat="server" Text='<%# Bind("IsUrgent") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Describes" SortExpression="item_desc">
                    <EditItemTemplate>
                        <asp:TextBox id="gv_txtDescribe" runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("item_desc") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                
                <asp:TemplateField HeaderText="Method" SortExpression="method">
                    <EditItemTemplate>
                        <asp:TextBox id="gv_txtMethod" runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# Bind("method") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Spec" SortExpression="spec">
                    <EditItemTemplate>
                        <asp:TextBox id="gv_txtSpec" runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label3" runat="server" Text='<%# Bind("spec") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Maxlimit" SortExpression="maxlimit">
                    <EditItemTemplate>
                        <asp:TextBox id="gv_txtMaxlimit" runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label4" runat="server" Text='<%# Bind("maxlimit") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Minlimit" SortExpression="minlimit">
                    <EditItemTemplate>
                        <asp:TextBox id="gv_txtMinlimit" runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label5" runat="server" Text='<%# Bind("minlimit") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Is Week Base" SortExpression="IsWeekly">
                    <EditItemTemplate>
                        <asp:RadioButton id="gv_rdoIsWeekly_Y" GroupName="C" Text="Yes" runat="server" 
                            AutoPostBack="True" oncheckedchanged="gv_rdoIsWeekly_Y_CheckedChanged"></asp:RadioButton>
                        <br />
                        <asp:RadioButton id="gv_rdoIsWeekly_N" GroupName="C" Text="No" runat="server" 
                            AutoPostBack="True" oncheckedchanged="gv_rdoIsWeekly_N_CheckedChanged"></asp:RadioButton>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label20" runat="server" Text='<%# Bind("IsWeekly") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="StartWeek" SortExpression="StartWeek">
                    <EditItemTemplate>
                        <asp:TextBox id="gv_txtStartWeek" Text='<%# Bind("StartWeek") %>' runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label21" runat="server" Text='<%# Bind("StartWeek") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Frequency" SortExpression="Frequency">
                    <EditItemTemplate>
                        <asp:TextBox id="gv_txtFrequency" Text='<%# Bind("Frequency") %>' runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label22" runat="server" Text='<%# Bind("Frequency") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Is Month Base" SortExpression="IsMonthly">
                    <EditItemTemplate>
                        <asp:RadioButton id="gv_rdoIsMonthly_Y" GroupName="F" Text="Yes" runat="server" 
                            AutoPostBack="True" oncheckedchanged="gv_rdoIsMonthly_Y_CheckedChanged"></asp:RadioButton>
                        <br />
                        <asp:RadioButton id="gv_rdoIsMonthly_N" GroupName="F" Text="No" runat="server" 
                            AutoPostBack="True" oncheckedchanged="gv_rdoIsMonthly_N_CheckedChanged"></asp:RadioButton>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label37" runat="server" Text='<%# Bind("IsMonthly") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="StartMonth" SortExpression="StartMonth">
                    <EditItemTemplate>
                        <asp:TextBox id="gv_txtStartMonth" Text='<%# Bind("StartMonth") %>' runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label38" runat="server" Text='<%# Bind("StartMonth") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Month_Frequency" SortExpression="Month_Frequency">
                    <EditItemTemplate>
                        <asp:TextBox id="gv_txtMonth_Frequency" Text='<%# Bind("Month_Frequency") %>' runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label39" runat="server" Text='<%# Bind("Month_Frequency") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>

                <asp:TemplateField HeaderText="Edit">
                    <EditItemTemplate>
                        <asp:LinkButton ID="lbUpdate" runat="server" CommandName="Update">Submit</asp:LinkButton>
                        <br />
                        <asp:LinkButton ID="lbCancelUpdate" runat="server" CommandName="Cancel">Cancel</asp:LinkButton>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:ImageButton ID="LinkButton1" CommandName="Edit" CommandArgument='<%# Eval("ItemID") %>' ImageUrl="../../images/icon_edit.png" runat="server" Width="18px" ToolTip="Edit"  />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Delete">
                    <ItemTemplate>                        
                        <asp:ImageButton ID="lbtnDelete" CommandName="myDelete" OnClientClick="return confirm('Are you sure of deleteing this sheet category?')"
                             CommandArgument='<%# Eval("ItemID") %>' ImageUrl="../../images/icon_delete.png" runat="server" Width="18px" ToolTip="Delete"  />
                    </ItemTemplate>   
                    <ItemStyle HorizontalAlign="Center" />          
                </asp:TemplateField> 
            </Columns>
            <PagerTemplate>
                <div style="text-align:center;" id="page">
                共<asp:Label ID="lblTotalCount" runat="server" Text=""></asp:Label>筆 │ 
                <asp:Label ID="lblPage" runat="server" ></asp:Label> / <asp:Label ID="lblTotalPage" runat="server" ></asp:Label>頁 │ 
                <asp:LinkButton ID="lbnFirst" runat="Server" Text="第一頁" onclick="lbnFirst_Click" ></asp:LinkButton> │ 
                <asp:LinkButton ID="lbnPrev" runat="server" Text="上一頁" onclick="lbnPrev_Click" ></asp:LinkButton> │ 
                <asp:LinkButton ID="lbnNext" runat="Server" Text="下一頁" onclick="lbnNext_Click"></asp:LinkButton> │ 
                <asp:LinkButton ID="lbnLast" runat="Server" Text="最後頁" onclick="lbnLast_Click" ></asp:LinkButton> │ 
                到第<asp:TextBox ID="inPageNum" Width="20px" runat="server"></asp:TextBox>頁： 
                每頁<asp:TextBox ID="txtSizePage" Width="25px" runat="server"></asp:TextBox>筆
                <asp:Button ID="btnGo" runat="server" Text="Go" onclick="btnGo_Click"/>
                <br />
                </div>
            </PagerTemplate>
            <FooterStyle BackColor="#CCCCCC" />
            <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="Black" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#CCCCCC" />  
        </asp:GridView>
        
        
    <asp:DetailsView ID="DetailsView1" runat="server"  AutoGenerateRows="False"
         CellPadding="4" ForeColor="#333333" GridLines="None" 
         OnItemCommand="DetailsView1_ItemCommand" 
         OnModeChanging="DetailsView1_ModeChanging" 
         OnItemInserting="DetailsView1_ItemInserting" OnDataBound="DetailsView1_DataBound">
                <Fields>
                <asp:TemplateField HeaderText="DataType">
                        <InsertItemTemplate>
                            <asp:DropDownList ID="dv_ddlDataType" runat="server" AutoPostBack="True" OnSelectedIndexChanged="dv_ddlDataType_SelectedIndexChanged" ></asp:DropDownList>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label1" runat="server" Text='<%# Bind("DataTypeID") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>   
                    <asp:TemplateField HeaderText="IsNumeric">
                        <InsertItemTemplate>
                            <asp:RadioButton id="dv_RdoIsNumeric_Y" Text="Yes" groupName="X" runat="server" AutoPostBack="True" OnCheckedChanged="dv_RdoIsNumeric_Y_CheckedChanged"></asp:RadioButton>
                            <asp:RadioButton id="dv_RdoIsNumeric_N" Text="No" groupName="X" runat="server" AutoPostBack="True" OnCheckedChanged="dv_RdoIsNumeric_N_CheckedChanged"></asp:RadioButton>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label9" runat="server" Text='<%# Bind("IsNumeric") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField> 
                    <asp:TemplateField HeaderText="IsUrgent">
                        <InsertItemTemplate>
                            <asp:RadioButton id="dv_RdoIsUrgent_Y" Text="Yes" groupName="Z" runat="server"></asp:RadioButton>
                            <asp:RadioButton id="dv_RdoIsUrgent_N" Text="No" groupName="Z" runat="server"></asp:RadioButton>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label8" runat="server" Text='<%# Bind("IsUrgent") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>                                     
                    <asp:TemplateField HeaderText="describes">
                        <InsertItemTemplate>
                            <asp:TextBox ID="dv_describes" runat="server" Text='<%# Bind("describes") %>' width="500px"></asp:TextBox>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label2" runat="server" Text='<%# Bind("describes") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Method">
                        <InsertItemTemplate>
                            <asp:TextBox ID="dv_Method" runat="server" Text='<%# Bind("Method") %>' width="500px"></asp:TextBox>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label3" runat="server" Text='<%# Bind("Method") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>     
                    <asp:TemplateField HeaderText="Spec">
                        <InsertItemTemplate>
                            <asp:TextBox ID="dv_Spec" runat="server" Text='<%# Bind("Spec") %>' width="500px"></asp:TextBox>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label4" runat="server" Text='<%# Bind("Spec") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>     
                    <asp:TemplateField HeaderText="MaxLimit">
                        <InsertItemTemplate>
                            <asp:TextBox ID="dv_MaxLimit" runat="server" Text='<%# Bind("MaxLimit") %>'></asp:TextBox>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label5" runat="server" Text='<%# Bind("MaxLimit") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="MinLimit">
                        <InsertItemTemplate>
                            <asp:TextBox ID="dv_MinLimit" runat="server" Text='<%# Bind("MinLimit") %>'></asp:TextBox>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label6" runat="server" Text='<%# Bind("MinLimit") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>  
                    <asp:TemplateField HeaderText="Is Week Base ">
                        <InsertItemTemplate>
                            <asp:RadioButton id="dv_rdoIsWeekly_Y" Text="Yes" groupName="D" runat="server" AutoPostBack="true" OnCheckedChanged="dv_rdoIsWeekly_Y_CheckedChanged"></asp:RadioButton>
                            <asp:RadioButton id="dv_rdoIsWeekly_N" Text="No" groupName="D" runat="server" AutoPostBack="true" OnCheckedChanged="dv_rdoIsWeekly_N_CheckedChanged"></asp:RadioButton>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label30" runat="server" Text='<%# Bind("IsWeekly") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Start Week">
                        <InsertItemTemplate>
                            <asp:TextBox ID="dv_StartWeek" runat="server" Text='<%# Bind("StartWeek") %>'></asp:TextBox>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label31" runat="server" Text='<%# Bind("StartWeek") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Week Frequency">
                        <InsertItemTemplate>
                            <asp:TextBox ID="dv_Frequency" runat="server" Text='<%# Bind("Frequency") %>'></asp:TextBox>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label32" runat="server" Text='<%# Bind("Frequency") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>

                    <asp:TemplateField HeaderText="Is Month Base">
                        <InsertItemTemplate>
                            <asp:RadioButton id="dv_rdoIsMonthly_Y" Text="Yes" groupName="E" runat="server" AutoPostBack="true" OnCheckedChanged="dv_rdoIsMonthly_Y_CheckedChanged"></asp:RadioButton>
                            <asp:RadioButton id="dv_rdoIsMonthly_N" Text="No" groupName="E" runat="server" AutoPostBack="true" OnCheckedChanged="dv_rdoIsMonthly_N_CheckedChanged"></asp:RadioButton>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label33" runat="server" Text='<%# Bind("IsMonthly") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Start Month">
                        <InsertItemTemplate>
                            <asp:TextBox ID="dv_StartMonth" runat="server" Text='<%# Bind("StartMonth") %>'></asp:TextBox>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label34" runat="server" Text='<%# Bind("StartMonth") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Month Frequency">
                        <InsertItemTemplate>
                            <asp:TextBox ID="dv_Month_Frequency" runat="server" Text='<%# Bind("Month_Frequency") %>'></asp:TextBox>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label35" runat="server" Text='<%# Bind("Month_Frequency") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>

                    <asp:TemplateField HeaderText="Action" ShowHeader="False">
                        <InsertItemTemplate>
                            <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="True" CommandName="Insert"
                                Text="Insert"></asp:LinkButton>
                            <asp:LinkButton ID="LinkButton2" runat="server" CausesValidation="False" CommandName="Cancel"
                                Text="Cancel"></asp:LinkButton>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="False" CommandName="New"
                                Text="新增"></asp:LinkButton>
                        </ItemTemplate>
                    </asp:TemplateField>
                </Fields>
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <CommandRowStyle BackColor="#E2DED6" Font-Bold="True" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <FieldHeaderStyle BackColor="#E9ECF1" Font-Bold="True" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <EditRowStyle BackColor="#999999" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
        </asp:DetailsView>
        
    </ContentTemplate>
</asp:UpdatePanel>
</asp:Content>
