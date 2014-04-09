<%@ Page Language="C#" AutoEventWireup="true" CodeFile="UpdateLog.aspx.cs" Inherits="EPM_Web.Alan.UpdateLog" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Log</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:GridView ID="GridView1" runat="server" BorderWidth="1px" BorderStyle="Solid"
                BorderColor="#999999" BackColor="White"                 
                GridLines="Vertical" ForeColor="Black" CellPadding="3" AutoGenerateColumns="False" EmptyDataText="無修改Log">
                <Columns>
                    <asp:TemplateField HeaderText="No">
                        <ItemTemplate>
                            <%# (Container.DataItemIndex+1).ToString()%>
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:TemplateField>                 
                    <asp:BoundField DataField="item_desc" HeaderText="item_desc" ></asp:BoundField>
                    <asp:BoundField DataField="OldValue" HeaderText="Old Value" ></asp:BoundField>
                    <asp:BoundField DataField="newValue" HeaderText="New Value" ></asp:BoundField>
                    <asp:BoundField DataField="account" HeaderText="Update Account" ></asp:BoundField>
                    <asp:BoundField DataField="UpdateTime" HeaderText="Update Time" ></asp:BoundField>                    
                </Columns>
                <FooterStyle BackColor="#CCCCCC"></FooterStyle>
                
                <PagerStyle HorizontalAlign="Center" BackColor="#999999" ForeColor="Black"></PagerStyle>
                <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White"></SelectedRowStyle>
                <HeaderStyle BackColor="Black" Font-Bold="True" ForeColor="White"></HeaderStyle>
                <AlternatingRowStyle BackColor="#CCCCCC"></AlternatingRowStyle>
            </asp:GridView>
    </div>
    </form>    
    
</body>
</html>
