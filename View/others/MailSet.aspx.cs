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
using System.Collections.Generic;

namespace EPM_Web.Alan
{
    public partial class MailSet : System.Web.UI.Page
    {
        public static string HrConn = System.Configuration.ConfigurationManager.ConnectionStrings["HR"].ToString();
        public static string Conn = System.Configuration.ConfigurationManager.ConnectionStrings["EPM"].ToString();
        public Common.AdoDbConn adoHr;
        public Common.AdoDbConn ado;
        public static int chkColumns = 15;
        public static string sheet_categoryID;

        protected void Page_Init(object sender, EventArgs e)
        {
            Common.MyFunc func = new Common.MyFunc();
            func.checkLogin();
            sheet_categoryID = Request.QueryString["sheet_categoryID"];
            initial();  
        }
        protected void Page_Load(object sender, EventArgs e)
        {   }


        private void initial()
        {
            Common.MyFunc func = new Common.MyFunc();
            DataTable dtDept = new DataTable();
            dtDept = func.getEmpDeptDt();//EMP_department

            PH.Controls.Add(new LiteralControl("<ul id='tab' class='nav nav-tabs'>"));
            foreach (DataRow dr in dtDept.Rows)
            {
                if(dr == dtDept.Rows[0])
                    PH.Controls.Add(new LiteralControl(string.Format(@"<li class='active'><a href='#{0}' data-toggle='tab'>{0}</a></li>", dr["DEPT_NAME"].ToString().Replace("PA-PACKAGE SAW", "PA-PACKAGESAW"))));
                else
                    PH.Controls.Add(new LiteralControl(string.Format(@"<li><a href='#{0}' data-toggle='tab'>{0}</a></li>", dr["DEPT_NAME"].ToString().Replace("PA-PACKAGE SAW", "PA-PACKAGESAW"))));
            }
            PH.Controls.Add(new LiteralControl("</ul> <div id='myTabContent' class='tab-content'>"));

            for (int i = 0; i < dtDept.Rows.Count; i++)
            {                
                EMP_dynamic_create(dtDept.Rows[i]["DEPT_NAME"].ToString(), i); 
            }
            PH.Controls.Add(new LiteralControl("<br />"));

            Button btn = new Button();
            btn.ID = "btnSubmit";
            btn.OnClientClick = "return confirm('確定更新名單?');";
            btn.Text = "Submit";
            btn.CssClass = "btn btn-primary";
            btn.Click += new EventHandler(btnSubmit_Click);
            PH.Controls.Add(btn);

            PH.Controls.Add(new LiteralControl("</div> "));
        }
        //dynamic create checkbox
        private void EMP_dynamic_create(string deptName, int tabIndex)
        {
            Common.MyFunc func = new Common.MyFunc();
            List<string> empLst = new List<string>();
            DataTable dtEmp = new DataTable();

            Panel pnl = new Panel();
            if(tabIndex.Equals(0))
                PH.Controls.Add(new LiteralControl(string.Format(@"<div class='tab-pane fade in active' id='{0}'> ", deptName.Replace("PA-PACKAGE SAW", "PA-PACKAGESAW"))));
            else
                PH.Controls.Add(new LiteralControl(string.Format(@"<div class='tab-pane fade' id='{0}'> ", deptName.Replace("PA-PACKAGE SAW", "PA-PACKAGESAW"))));

            empLst = getCheckedEmpList(deptName);//selected items
            dtEmp = func.getDeptAllData(deptName);//all items

            pnl.Controls.Add(new LiteralControl("<table>"));
            for (int i = 0; i < dtEmp.Rows.Count; i++)
            {
                if ((i % chkColumns == 0) || i == 0) { pnl.Controls.Add(new LiteralControl("<tr>")); }
                CheckBox chk = new CheckBox();
                chk.Text = string.Format(@"{0}_{1}", dtEmp.Rows[i]["name"].ToString(),
                                                      dtEmp.Rows[i]["EMP_NO"].ToString());
                chk.Checked = empLst.Contains(dtEmp.Rows[i]["EMP_NO"].ToString());

                pnl.Controls.Add(new LiteralControl("<td>"));
                pnl.Controls.Add(chk);
                pnl.Controls.Add(new LiteralControl("</td>"));
                if ((i + 1) % chkColumns == 0) { pnl.Controls.Add(new LiteralControl("</tr>")); }
            }
            pnl.Controls.Add(new LiteralControl("</table>"));
            PH.Controls.Add(pnl);
            PH.Controls.Add(new LiteralControl("</div> "));
        }


        private List<string> getCheckedEmpList(string deptName)
        {
            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            string Str = string.Format(@"SELECT EMP_NO FROM vw_MailContact 
                                         WHERE dept=:dept AND sheet_categoryID=:sheet_categoryID");
            object[] para = new object[] { deptName, sheet_categoryID };
            DataTable dt = new DataTable();
            dt = ado.loadDataTable(Str, para, "vw_MailContact");
            List<string> lst = new List<string>();
            foreach (DataRow dr in dt.Rows) { lst.Add(dr["EMP_NO"].ToString()); }
            return lst;
        }

        protected void btnSubmit_Click(object sender, EventArgs e)
        {

            List<string[]> lst = new List<string[]>();

            Common.MyFunc func = new Common.MyFunc();
            PlaceHolder panel = (PlaceHolder)Page.FindControl("PH");
            foreach (object ctrl in ((PlaceHolder)panel).Controls)
            {
                if (ctrl.GetType().Name.ToString().Equals("Panel"))
                {
                    foreach (object obj in ((Panel)ctrl).Controls)
                    {
                        switch (obj.GetType().Name.ToString())
                        {
                            case "CheckBox":
                                if (((CheckBox)obj).Checked)
                                {
                                    string[] ary = ((CheckBox)obj).Text.Split('_');
                                    lst.Add(new string[] { ary[0], ary[1] });
                                }
                                break;
                        }
                    }
                }
            }
            

            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            #region delete old all
            string delstr = string.Format(@"Delete from MailContact WHERE sheet_categoryID='{0}'", sheet_categoryID);
            ado.dbNonQuery(delstr, null);
            #endregion
            #region insert new all from list
            for (int i = 0; i < lst.Count; i++)
            {
                DataTable dt = func.getPersonalData(lst[i][1]);
                string sqlstr = string.Format(@"INSERT INTO MailContact(MailContactID,sheet_categoryID,UsersID) 
                                                VALUES (MailContact_sequence.nextval,:sheet_categoryID,:UsersID)");
                object[] para = new object[] { sheet_categoryID, dt.Rows[0]["UsersID"].ToString() };
                ado.dbNonQuery(sqlstr, para);
            }

            #endregion

            string strMsg = "已成功更新名單";
            ScriptManager.RegisterStartupScript(this, this.GetType(), "", string.Format(@"alert('{0}');", strMsg), true);

        }
}
}