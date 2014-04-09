<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Login.aspx.cs" Inherits="EPM_Web.Alan.Login" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <script type="text/javascript" src="~/js/jquery-1.10.0.js"></script>    
    <%--bootstrap--%>
    <script type="text/javascript" src="~/js/bootstrap.min.js"></script>
    <link href="~/css/bootstrap-responsive.min.css" rel="stylesheet" type="text/css" />
    <link href="~/css/bootstrap.min.css" rel="stylesheet" type="text/css" />
    
    <title>ePM 登入系統</title>
    <style type="text/css">
    .pie{                
        border: 1px solid #696;
        font-family:Georgia;
        padding: 25px 5px;
        text-align: left; width: 100%;
        -webkit-border-radius: 40px;
        -moz-border-radius: 40px;
        border-radius: 20px;
        -webkit-box-shadow: #666 0px 2px 3px;
        -moz-box-shadow: #666 0px 2px 3px;
        box-shadow: #666 0px 2px 3px;
        background: #EEFF99;
        background: -webkit-gradient(linear, 0 0, 0 bottom, from(#EEFF99), to(#66EE33));
        background: -webkit-linear-gradient(#EEFF99, #66EE33);
        background: -moz-linear-gradient(#EEFF99, #66EE33);
        background: -ms-linear-gradient(#EEFF99, #66EE33);
        background: -o-linear-gradient(#EEFF99, #66EE33);
        background: linear-gradient(#EEFF99, #66EE33);
        -pie-background: linear-gradient(#EEFF99, #66EE33);
        behavior: url(/pie/PIE.htc);
    }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    
    <table>
        <tr>
            <td style="height:70px; text-align:left">
                <asp:Image ID="Image4" ImageUrl="../../images/bg_logo.gif" runat="server"/>
            </td>
            <td style="text-align:left">
                <asp:Image ID="Image2" ImageUrl="../../images/banner_ePM1.png" runat="server"/> 
            </td>
        </tr>
        <tr><td colspan="2"><hr style=" background-color:#DDDDDD; height:1px;" /></td></tr>
        <tr>
            <td style="vertical-align:top;">
                
                <table>  
                    <tr>
                        <td style="text-align:right; vertical-align:top;">Sheet Category：</td>
                        <td>
                            <asp:DropDownList ID="ddl1" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddl1_SelectedIndexChanged" Width="100%"></asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td style="text-align:right; vertical-align:top;">Tester：</td>
                        <td>
                            <asp:DropDownList ID="ddl2" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddl2_SelectedIndexChanged">
                            </asp:DropDownList>
                        </td>
                    </tr>                  
                    <tr>
                        <td style="text-align:right; vertical-align:top;">帳號(Account)：</td>
                        <td><asp:TextBox ID="txtAcc" runat="server" EnableViewState="False"></asp:TextBox></td>  
                                 
                    </tr>
                    <tr>
                        <td style="text-align:right;">密碼(Password)：</td>        
                        <td><asp:TextBox ID="txtPw" runat="server" TextMode="Password"></asp:TextBox></td>            
                    </tr>
                    <tr>
                        <td colspan="2" style="text-align:right">
                            <asp:Button ID="btnSubmit" runat="server" Text="Login" CssClass="btn btn-primary" OnClick="btnSubmit_Click" />
                            <input type="button" value="Close" class="btn btn-inverse" onclick="window.open('','_self','');window.close();"/>
                        </td>
                    </tr>
                </table>
            </td>
            <td style="padding:0 0 0 20px;">
                <div class="pie">
                    The web page you tried to access is protected. In order to access it you need to fill<br />
                    in your Username and Password and click on the "log In" button.<br />
                    請輸入您的薪號及密碼，然後按Log In。<br />
                    <br />
                    As username type in your payroll no.<br />
                    請輸入薪號7碼。<br />
                    e.g. 1234567 (7 digits)<br />
                    <br />
                    As default password type in your the last 4 digits of ID / passport no.<br />
                    預設密碼請輸入<div style="color:Red; display:inline;"><b>身分證號碼後四碼</b></div>。<br />
                    e.g. 1234 (4 digits)<br />
                    <br />
                    Owner：T.W.Lin #8803<br />
                    IT Support : Alan Kuo #8236
                    <br />
                    <div style="color:red; display:inline;">Download：</div><asp:HyperLink ID="HyperLink1" NavigateUrl="http://165.114.64.94/ePM_prod/download/ePM%20User%20Manual.pptx" runat="server">User Manual</asp:HyperLink>
                </div>
            </td>
        </tr>
    </table>
    
    </form>    
</body>
</html>
