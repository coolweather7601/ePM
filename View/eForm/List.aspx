<%@ Page Language="C#" MasterPageFile="~/View/Master/MasterPage.master" AutoEventWireup="true" CodeFile="List.aspx.cs" Inherits="EPM_Web.Alan.List" Title="ePM" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>

<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>
    
    
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder2" Runat="Server">
    <asp:UpdatePanel id="up" runat="server">
    <ContentTemplate>
    <h2>List</h2>
        
    <%-- Tab--%>
    <table>
        <tbody>
            <tr>
                <td style="text-align:left">Sheet</td>
                <td colspan="4">
                    ：<asp:DropDownList id="ddl_sheetCategory" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddl_sheetCategory_SelectedIndexChanged"></asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td style="text-align:left">Tester</td>
                <td>：<asp:TextBox id="txtTester" runat="server"></asp:TextBox></td>
                <td style="text-align:left">Location</td>
                <td>：<asp:TextBox id="txtLocation" runat="server"></asp:TextBox></td>
                <td>Owner：<asp:TextBox id="txtOwner" runat="server"></asp:TextBox></td>
            </tr>
            <tr>
                <td style="text-align:left">Start Week</td>
                <td>：<asp:TextBox id="txtStartdate" runat="server"></asp:TextBox></td>
                <td style="text-align:left">End Week</td>
                <td>：<asp:TextBox id="txtEnddate" runat="server"></asp:TextBox></td>
                <td style="text-align:right"><asp:Button ID="btnSearch" runat="server" Text="Search" OnClick="btnSearch_Click"></asp:Button></td>
            </tr>
        </tbody>
    </table>
    <%--TabContent--%>
    <h3><%=title_sheet_category %></h3>
    <p>
        <asp:GridView ID="GridViewList" runat="server" BorderWidth="1px" BorderStyle="Solid"
            BorderColor="#999999" BackColor="White" OnPageIndexChanging="GridViewList_PageIndexChanging"
            AllowPaging="True" OnRowDataBound="GridViewList_RowDataBound" OnRowCancelingEdit="GridViewList_RowCancelingEdit"
            OnRowUpdating="GridViewList_RowUpdating" OnRowEditing="GridViewList_RowEditing" OnRowCommand="GridViewList_RowCommand"
            GridLines="Vertical" ForeColor="Black" CellPadding="3" AutoGenerateColumns="False"
            DataKeyNames="SheetID">
            <Columns>
                <asp:TemplateField HeaderText="No">
                    <ItemTemplate>
                        <%# (Container.DataItemIndex+1).ToString()%>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                </asp:TemplateField>                
                <asp:BoundField DataField="sheetCategory_desc" HeaderText="sheetCategory_desc" ReadOnly="True"></asp:BoundField>
                <asp:BoundField DataField="machine" HeaderText="Tester" ReadOnly="True"></asp:BoundField>
                <asp:BoundField DataField="location" HeaderText="Location" ReadOnly="True"></asp:BoundField>
                <asp:BoundField DataField="weekID" HeaderText="weekID" ReadOnly="True"></asp:BoundField>
                <asp:BoundField DataField="insertTime" HeaderText="Time" ReadOnly="True"></asp:BoundField>
                <asp:TemplateField HeaderText="Owner">
                    <ItemTemplate>
                        <%--<asp:Label runat="server" Text='<%# Bind("machine") %>' ID="gv_lblOwner"></asp:Label>--%>
                        <asp:Label runat="server" Text='<%# Bind("owner1") %>' ID="gv_lblOwner"></asp:Label>
                        <asp:Label runat="server" Text='<%# Bind("owner2") %>' ID="Label1"></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>

                <asp:TemplateField HeaderText="memo">
                    <EditItemTemplate>
                        <asp:TextBox ID="gv_txtMemo" Text='<%# Bind("memo") %>' Rows="7" Width="300px" runat="server" 
                            TextMode="MultiLine"></asp:TextBox>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label runat="server" Text='<%# Bind("memo") %>' ID="Label0"></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Log" Visible="true">
                    <ItemTemplate>
                        <asp:ImageButton ID="btnLog1" CommandName="log" CommandArgument='<%# Eval("SheetID") %>'
                            ImageUrl="../../images/icon_log.png" runat="server" Width="18px" ToolTip="Log"  />
                    </ItemTemplate>
                </asp:TemplateField>                    
                <asp:TemplateField HeaderText="Edit" Visible="true">
                    <ItemTemplate>
                        <asp:ImageButton ID="btn1" CommandName="view" CommandArgument='<%# Eval("SheetID") %>'
                            ImageUrl="../../images/icon_edit.png" runat="server" Width="18px" ToolTip="Edit"  />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Memo" Visible="true">
                    <EditItemTemplate>
                        <asp:LinkButton ID="lbUpdate" runat="server" CommandName="Update">Submit</asp:LinkButton>
                        |
                        <asp:LinkButton ID="lbCancelUpdate" runat="server" CommandName="Cancel">Cancel</asp:LinkButton>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:ImageButton ID="btn17" CommandName="Edit" CommandArgument='<%# Eval("SheetID") %>'
                            ImageUrl="../../images/icon_fix.ico" runat="server" Width="18px" ToolTip="Fix"  />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Delete" Visible="false">
                    <ItemTemplate>
                        <asp:ImageButton ID="btn2" CommandName="del" OnClientClick="return confirm('Are you sure of deleting this sheet?');"
                            CommandArgument='<%# Eval("SheetID") %>' ImageUrl="../../images/icon_delete.png"
                            runat="server" Width="18px" ToolTip="Delete"  />
                    </ItemTemplate>
                </asp:TemplateField>
                
            </Columns>
            <FooterStyle BackColor="#CCCCCC"></FooterStyle>
            <PagerTemplate>
                <div style="text-align: center;" id="page">
                    共<asp:Label ID="lblTotalCount" runat="server" Text=""></asp:Label>筆 │
                    <asp:Label ID="lblPage" runat="server"></asp:Label>
                    /
                    <asp:Label ID="lblTotalPage" runat="server"></asp:Label>頁 │
                    <asp:LinkButton ID="lbnFirst" runat="Server" Text="第一頁" OnClick="lbnFirst_Click"></asp:LinkButton>
                    │
                    <asp:LinkButton ID="lbnPrev" runat="server" Text="上一頁" OnClick="lbnPrev_Click"></asp:LinkButton>
                    │
                    <asp:LinkButton ID="lbnNext" runat="Server" Text="下一頁" OnClick="lbnNext_Click"></asp:LinkButton>
                    │
                    <asp:LinkButton ID="lbnLast" runat="Server" Text="最後頁" OnClick="lbnLast_Click"></asp:LinkButton>
                    │ 到第 <asp:TextBox onKeyDown="preventTextEnterEvent();" ID="inPageNum" Width="20px" runat="server"></asp:TextBox>頁： 每頁 <asp:TextBox onKeyDown="preventTextEnterEvent();"
                        ID="txtSizePage" Width="25px" runat="server"></asp:TextBox>筆
                    <asp:Button ID="btnGo" runat="server" Text="Go" OnClick="btnGo_Click" />
                    <br />
                </div>
            </PagerTemplate>
            <PagerStyle HorizontalAlign="Center" BackColor="#999999" ForeColor="Black"></PagerStyle>
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White"></SelectedRowStyle>
            <HeaderStyle BackColor="Black" Font-Bold="True" ForeColor="White"></HeaderStyle>
            <AlternatingRowStyle BackColor="#CCCCCC"></AlternatingRowStyle>
        </asp:GridView>
    </p>
    
    </ContentTemplate>
</asp:UpdatePanel>
</asp:Content>

