<%@ Page Language="C#" MasterPageFile="~/View/Master/MasterPage.master" AutoEventWireup="true" CodeFile="ACS_Manage.aspx.cs" Inherits="EPM_Web.Alan.ACS_Manage" Title="ePM" %>
<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>

<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder2" Runat="Server">

<asp:UpdatePanel id="up" runat="server">
    <ContentTemplate>
    
    <h2>ACS_Manage</h2>
       
    <table>
        <tr>
            <td>Tester：</td>
            <td><asp:TextBox ID="txtTester" runat="server"></asp:TextBox></td>
            <td>Location：</td>
            <td><asp:TextBox ID="txtLocation" runat="server"></asp:TextBox></td>
            <td style="text-align:right;"><asp:Button ID="btnSearch" runat="server" Text="Search" OnClick="btnSearch_Click" /></td>
        </tr>
        <tr>    
            <td colspan="5">
                <asp:Button ID="btnNew" runat="server" Visible="false" Text="New Data" OnClick="btnNew_Click" />
                <asp:Button ID="btnImport" runat="server" Visible="False" Text="Import" OnClick="btnImport_Click" />
            </td>
        </tr>       
    </table> 
      
    <asp:GridView ID="GridViewAM" runat="server" DataKeyNames="ACS_ManageID" AutoGenerateColumns="False"
                      CellPadding="3" ForeColor="Black" GridLines="Vertical" AllowSorting="True" AllowPaging="True"
            OnSorting="GridViewAM_Sorting" 
            OnRowCancelingEdit="GridViewAM_RowCancelingEdit" 
            OnRowEditing="GridViewAM_RowEditing" 
            OnRowUpdating="GridViewAM_RowUpdating" 
            OnRowDataBound="GridViewAM_RowDataBound"  
            OnPageIndexChanging="GridViewAM_PageIndexChanging" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" OnRowCommand="GridViewAM_RowCommand" Width="65%">
            <Columns>                
                <asp:TemplateField HeaderText="No.">
                        <ItemTemplate>
                            <%# (Container.DataItemIndex +1).ToString()%>
                        </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="5%" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Tester" SortExpression="Tester">
                    <EditItemTemplate>
                        <asp:TextBox id="gv_txtTester" Text='<%# Bind("Tester") %>' runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("Tester") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle Width="30%" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Location" SortExpression="Location">
                    <EditItemTemplate>
                        <asp:TextBox id="gv_txtLocation" Text='<%# Bind("Location") %>' runat="server"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# Bind("Location") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle Width="30%" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="isSealing" SortExpression="isSealing">
                    <EditItemTemplate>
                        <asp:RadioButton ID="gv_rdoIsSealing_Y" Text="Yes" GroupName="AA" runat="server" />
                        <asp:RadioButton ID="gv_rdoIsSealing_N" Text="No" GroupName="AA" runat="server" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label21" runat="server" Text='<%# Bind("isSealing") %>'></asp:Label>
                    </ItemTemplate>
                    <HeaderStyle Width="10%" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Edit" Visible="false">
                    <EditItemTemplate>
                        <asp:LinkButton ID="lbUpdate" runat="server" CommandName="Update">Submit</asp:LinkButton>
                        <br />
                        <asp:LinkButton ID="lbCancelUpdate" runat="server" CommandName="Cancel">Cancel</asp:LinkButton>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:ImageButton ID="btn1" CommandName="Edit" CommandArgument='<%# Eval("ACS_ManageID") %>'
                                ImageUrl="../../images/icon_edit.png" runat="server" Width="18px" ToolTip="Edit"  />                        
                    </ItemTemplate>
                    <HeaderStyle Width="5%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Delete" Visible="false">
                    <ItemTemplate>                        
                        <asp:ImageButton ID="btn2" CommandName="myDelete" OnClientClick="return confirm('Are you sure of deleteing this tester?');"
                                CommandArgument='<%# Eval("ACS_ManageID") %>' ImageUrl="../../images/icon_delete.png"
                                runat="server" Width="18px" ToolTip="Delete"  />
                    </ItemTemplate>             
                    <HeaderStyle Width="5%" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField> 
                <asp:TemplateField HeaderText="Select" visible="false">
                    <ItemTemplate>                        
                        <asp:Button ID="gv_btnSelect" CommandName="mySelect" CommandArgument='<%# Eval("ACS_ManageID") %>' runat="server" Text="Button"></asp:Button>
                    </ItemTemplate>             
                    <HeaderStyle Width="10%" />
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
        
        <asp:DetailsView ID="DetailsView1" runat="server" Height="50px" Width="125px" AutoGenerateRows="False"
         CellPadding="4" ForeColor="#333333" GridLines="None" 
         OnItemCommand="DetailsView1_ItemCommand" 
         OnModeChanging="DetailsView1_ModeChanging" 
         OnItemInserting="DetailsView1_ItemInserting" OnDataBound="DetailsView1_DataBound">
                <Fields>                
                    <asp:TemplateField HeaderText="Tester">
                        <InsertItemTemplate>
                            <asp:TextBox ID="dv_txtTester" runat="server" Text='<%# Bind("Tester") %>'></asp:TextBox>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label3" runat="server" Text='<%# Bind("Tester") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>     
                    <asp:TemplateField HeaderText="Location">
                        <InsertItemTemplate>
                            <asp:TextBox ID="dv_txtLocation" runat="server" Text='<%# Bind("Location") %>'></asp:TextBox>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label4" runat="server" Text='<%# Bind("Location") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>  
                    <asp:TemplateField HeaderText="isSealing">
                        <InsertItemTemplate>
                            <asp:RadioButton ID="dv_rdoIsSealing_Y" Text="Yes" GroupName="BB" runat="server"></asp:RadioButton>
                            <asp:RadioButton ID="dv_rdoIsSealing_N" Text="No" GroupName="BB" runat="server"></asp:RadioButton>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label4" runat="server" Text='<%# Bind("isSealing") %>'></asp:Label>
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