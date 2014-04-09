<%@ Page Language="C#" AutoEventWireup="true" CodeFile="MailSet.aspx.cs" Inherits="EPM_Web.Alan.MailSet" %>

<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>MailSet</title>
    
    <script type="text/javascript" src="../../js/jquery-1.10.0.js"></script>
    
    <%--bootstrap--%>
    <script type="text/javascript" src="../../js/bootstrap.min.js"></script>
    <link href="../../css/bootstrap-responsive.min.css" rel="stylesheet" type="text/css" />
    <link href="../../css/bootstrap.min.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <ContentTemplate>            
                <h3>緊急聯絡人</h3>
                <asp:PlaceHolder id="PH" runat="server"></asp:PlaceHolder>      
            </ContentTemplate>
        </asp:UpdatePanel>
        
    </div>
    </form>
</body>
</html>
