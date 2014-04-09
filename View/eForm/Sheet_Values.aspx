<%@ Page Language="C#" MasterPageFile="~/View/Master/MasterPage.master" AutoEventWireup="true" CodeFile="Sheet_Values.aspx.cs" Inherits="EPM_Web.Alan.Sheet_Values" Title="ePM" %>
<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>
    
    
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder2" Runat="Server">
    <asp:UpdatePanel id="up" runat="server">
    <ContentTemplate>
    <h2><%=title_sheetCategory_desc %></h2>
    <table id="autoTable">        
            <tbody>
                <tr>
                    <td style="text-align: right">
                        機台編號 Machine No.：</td>
                    <td>
                        <asp:TextBox onkeydown="preventTextEnterEvent();" ID="txtMachine" runat="server"
                            CssClass="input-large" Enabled="false"></asp:TextBox></td>
                    <td style="text-align: right">
                        機台位置 Location.：</td>
                    <td>
                        <asp:TextBox onkeydown="preventTextEnterEvent();" ID="txtLocation" runat="server"
                            CssClass="input-large" Enabled="false"></asp:TextBox></td>
                </tr>
                <tr>
                    <td style="text-align: right">
                        基準/週別 Time/Week：</td>
                    <td>
                        <asp:TextBox onkeydown="preventTextEnterEvent();" ID="txtWeek" runat="server"
                            CssClass="input-large" Enabled="false"></asp:TextBox></td>
                    <td style="text-align: right">
                        操作人員 Operator：</td>
                    <td>
                        <asp:TextBox onkeydown="preventTextEnterEvent();" ID="txtOperator" runat="server"
                            CssClass="input-large" Enabled="false"></asp:TextBox></td>
                </tr>
            </tbody>
        </table>
        
    <asp:PlaceHolder id="PH" runat="server"></asp:PlaceHolder>
    </ContentTemplate>
</asp:UpdatePanel>
</asp:Content>
