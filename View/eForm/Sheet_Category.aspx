<%@ Page Language="C#" MasterPageFile="~/View/Master/MasterPage.master" AutoEventWireup="true" CodeFile="Sheet_Category.aspx.cs" Inherits="EPM_Web.Alan.Sheet_Category" Title="ePM" %>
<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>
    
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder2" Runat="Server">

<asp:UpdatePanel id="up" runat="server">
    <ContentTemplate>
    
    <h2>Sheet_Category</h2>
    <table>
        <%--<tr>
            <td>DocNumber：</td>
            <td colspan="4"><asp:TextBox ID="txtDocNumber" runat="server"></asp:TextBox></td>
        </tr>
        <tr>
            <td>Describe：</td>
            <td><asp:TextBox ID="txtDescribe" runat="server"></asp:TextBox></td>
            <td>Tester：</td>
            <td><asp:TextBox ID="txtTester" runat="server"></asp:TextBox></td>
            <td><asp:Button ID="btnSearch" runat="server" Text="Search" OnClick="btnSearch_Click" /></td>
        </tr>--%>
        <tr>
            <td>DocNumber：</td>
            <td><asp:TextBox ID="txtDocNumber" runat="server"></asp:TextBox></td>
            <td>Describe：</td>
            <td><asp:TextBox ID="txtDescribe" runat="server"></asp:TextBox></td>
            <td><asp:Button ID="btnSearch" runat="server" Text="Search" OnClick="btnSearch_Click" /></td>
        </tr>
        <tr>
            <td colspan="5"><asp:Button ID="btnNew" runat="server" Visible="false" Text="New Data" OnClick="btnNew_Click" /></td>
        </tr>
    </table>
                  
    <asp:GridView ID="GridViewSC" runat="server" DataKeyNames="Sheet_categoryID" AutoGenerateColumns="False" width="85%" 
                      CellPadding="3" ForeColor="Black" GridLines="Vertical" AllowSorting="True" AllowPaging="True"
            OnSorting="GridViewSC_Sorting" 
            OnRowCancelingEdit="GridViewSC_RowCancelingEdit" 
            OnRowEditing="GridViewSC_RowEditing" 
            OnRowUpdating="GridViewSC_RowUpdating" 
            OnRowDataBound="GridViewSC_RowDataBound"  
            OnPageIndexChanging="GridViewSC_PageIndexChanging" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" OnRowCommand="GridViewSC_RowCommand">            
            
            <Columns>                            
                <asp:TemplateField HeaderText="No.">
                        <ItemTemplate>
                            <%# (Container.DataItemIndex +1).ToString()%>
                        </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="3%" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="DocNumber" SortExpression="DocNumber">
                    <EditItemTemplate>
                        <asp:TextBox id="gv_txtDocNumber" Text='<%# Bind("DocNumber") %>' width="90%" runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label8" runat="server" Text='<%# Bind("DocNumber") %>'></asp:Label>
                    </ItemTemplate>                     
                    <HeaderStyle Width="12%" />
                </asp:TemplateField> 
                <asp:TemplateField HeaderText="Describes" SortExpression="describes">
                    <EditItemTemplate>
                        <asp:TextBox id="gv_txtDescribe" Text='<%# Bind("describes") %>' width="90%" runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("describes") %>'></asp:Label>
                    </ItemTemplate>                     
                    <HeaderStyle Width="29%" />
                </asp:TemplateField>   
                <asp:TemplateField HeaderText="Tester" SortExpression="Tester">
                    <ItemTemplate>
                        <asp:Label ID="gv_Tester" runat="server"></asp:Label>
                    </ItemTemplate>                     
                    <HeaderStyle Width="31%" />
                </asp:TemplateField> 
                <asp:TemplateField HeaderText="Edit" Visible="false">
                    <EditItemTemplate>
                        <asp:LinkButton ID="lbUpdate" runat="server" CommandName="Update">Submit</asp:LinkButton>
                        <br />
                        <asp:LinkButton ID="lbCancelUpdate" runat="server" CommandName="Cancel">Cancel</asp:LinkButton>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:ImageButton ID="btn1" CommandName="Edit" CommandArgument='<%# Eval("Sheet_categoryID") %>'
                                ImageUrl="../../images/icon_edit.png" runat="server" Width="18px" ToolTip="Edit"  />                                        
                    </ItemTemplate>                    
                    <HeaderStyle Width="3%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                 <asp:TemplateField HeaderText="Delete" Visible="false">
                    <ItemTemplate>
                        <asp:ImageButton ID="btn3" CommandName="myDelete" CommandArgument='<%# Eval("Sheet_categoryID") %>' onclientclick="return confirm('Are you sure of deleteing this sheet category?')"
                                ImageUrl="../../images/icon_delete.png" runat="server" Width="18px" ToolTip="Delete"  />
                    </ItemTemplate>                    
                    <HeaderStyle Width="3%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="View" Visible="false">
                    <ItemTemplate>                      
                        <asp:HyperLink ID="HyperLink1" ImageUrl="../../images/icon_preview.png" Width="20px" ToolTip="Preview" Target="_blank" NavigateUrl='<%# "Sheet_Values.aspx?sheet_categoryID="+Eval("Sheet_categoryID")+"&preview=Yes"%>' runat="server">Preview</asp:HyperLink>                                                     
                    </ItemTemplate>                    
                    <HeaderStyle Width="3%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>   
                <asp:TemplateField HeaderText="Copy" Visible="false">
                    <ItemTemplate>
                        <asp:ImageButton ID="btn2" CommandName="Copy" CommandArgument='<%# Eval("Sheet_categoryID") %>' onclientclick="return confirm('Are you sure of copying this sheet category?')"
                                ImageUrl="../../images/icon_copy.png" runat="server" Width="18px" ToolTip="Copy sheet"  />
                    </ItemTemplate>                    
                    <HeaderStyle Width="3%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>            
                <asp:TemplateField HeaderText="Items" Visible="false">
                    <ItemTemplate>
                        <asp:HyperLink ID="HyperLink2" ImageUrl="../../images/icon_item.png" Width="18px"  ToolTip="Items set" Target="_blank" NavigateUrl='<%# "item.aspx?sheet_categoryID="+Eval("Sheet_categoryID")%>' runat="server">Set</asp:HyperLink>
                    </ItemTemplate>                    
                    <HeaderStyle Width="3%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>                       
                <asp:TemplateField HeaderText="Bind" Visible="false">
                    <ItemTemplate>
                        <asp:HyperLink ID="HyperLink5" ImageUrl="../../images/icon_set.png" Width="18px"  ToolTip="Bind ACS testers" Target="_blank" NavigateUrl='<%# "../others/ACS_Manage.aspx?sheet_categoryID="+Eval("Sheet_categoryID")%>' runat="server">Bind</asp:HyperLink>
                    </ItemTemplate>                    
                    <HeaderStyle Width="3%" />
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
         OnItemInserting="DetailsView1_ItemInserting">
                <Fields>
                    <asp:TemplateField HeaderText="Describes">
                        <InsertItemTemplate>
                            <asp:TextBox ID="dv_describes" Width="400px" runat="server" Text='<%# Bind("describes") %>'></asp:TextBox>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label1" runat="server" Text='<%# Bind("describes") %>'></asp:Label>
                        </ItemTemplate>               
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="DocNumber">
                        <InsertItemTemplate>
                            <asp:TextBox ID="dv_DocNumber" Width="200px" runat="server" Text='<%# Bind("DocNumber") %>'></asp:TextBox>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label2" runat="server" Text='<%# Bind("DocNumber") %>'></asp:Label>
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
                            <asp:LinkButton ID="LinkButton3" runat="server" CausesValidation="False" CommandName="New"
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