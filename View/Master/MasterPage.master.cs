using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

public partial class MasterPage : System.Web.UI.MasterPage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        EPM_Web.Alan.Common.MyFunc func = new EPM_Web.Alan.Common.MyFunc();
        func.checkLogin();
        lblName.Text = string.Format(@"[{0}]", Session["Name"].ToString());
        
    }
    protected void btnLogout_Click(object sender, EventArgs e)
    {
        Session["account"] = null;
        Session["UserGroup"] = null;
        Session["Name"] = null;
        Session["Dept_Name"] = null;
        Response.Redirect("../others/Login.aspx");        
    }
    protected void btnRole_Click(object sender, EventArgs e)
    {
        Response.Redirect("../others/RoleManage.aspx");
    }
}
