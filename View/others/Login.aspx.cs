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

namespace EPM_Web.Alan
{
    public partial class Login : System.Web.UI.Page
    {
        public static string HrConn = System.Configuration.ConfigurationManager.ConnectionStrings["HR"].ToString();
        public static string Conn = System.Configuration.ConfigurationManager.ConnectionStrings["EPM"].ToString();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                ddl_bind();
                if (!string.IsNullOrEmpty(Request.QueryString["tester"]))
                {
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('機台號碼：{0}')", Request.QueryString["tester"]), true);
                }

                if (Session["account"] != null)
                {
                    string sheet_categoryID = Request.QueryString["sheet_categoryID"];
                    string sheetID = Request.QueryString["SheetID"];

                    Session["tester"] = Request.QueryString["tester"];
                    Session["conn"] = Request.QueryString["conn"];
                    PageSwitch(Request.QueryString["tester"], sheet_categoryID, sheetID);
                }
            }
        }
        protected void btnSubmit_Click(object sender, EventArgs e)
        {
            string sheet_categoryID = Request.QueryString["sheet_categoryID"];
            string sheetID = Request.QueryString["SheetID"];

            //ACS
            Session["tester"] = Request.QueryString["tester"];
            Session["conn"] = Request.QueryString["conn"];
            
            #region EPM
            Common.AdoDbConn ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            DataTable dt = new DataTable();
            string sqlstr = @"SELECT UsersID,RoleID,account,password,name,dept,emp_no,mail 
                              FROM Users 
                              WHERE account=:account AND password=:password
                              Order by usersid desc";
            object[] para = new object[] { txtAcc.Text.Trim(), txtPw.Text.Trim() };

            dt = ado.loadDataTable(sqlstr, para, "Users");
            if (dt.Rows.Count > 0)
            {
                Session["UsersID"] = dt.Rows[0]["UsersID"].ToString();
                Session["account"] = dt.Rows[0]["account"].ToString();
                Session["RoleID"] = dt.Rows[0]["RoleID"].ToString();
                Session["Name"] = dt.Rows[0]["name"].ToString();
                Session["Dept_Name"] = dt.Rows[0]["Dept"].ToString();

                PageSwitch(Request.QueryString["tester"], sheet_categoryID, sheetID);
                return;
            }
            #endregion

            #region EMP
            Common.AdoDbConn adoHr = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, HrConn);
            DataTable dtEmp = new DataTable();
            string strEmp = string.Format(@"SELECT ltrim(EMP_NO,'0') as EMP_NO,Emp_Name,Dept_NAME,passwd,email
                                            FROM   EMP_ACCESS_LIST 
                                            WHERE  Emp_No like '%{0}%' AND Passwd='{1}'", txtAcc.Text.Trim(), txtPw.Text.Trim());
            dtEmp = adoHr.loadDataTable(strEmp, null, "EMP_ACCESS_LIST");
            if (dtEmp.Rows.Count > 0)
            {
                Session["account"] = dtEmp.Rows[0]["Emp_No"].ToString();
                Session["Name"] = dtEmp.Rows[0]["Emp_Name"].ToString();
                Session["Dept_Name"] = dtEmp.Rows[0]["Dept_Name"].ToString();
                Session["RoleID"] = "2";//general user

                string insertStr = @"INSERT INTO Users(UsersID,RoleID,account,password,name,dept,emp_no,mail )
                                     VALUES (Users_sequence.nextval,'2',:account,:password,:name,:dept,:emp_no,:mail)";
                object[] param = new object[] { dtEmp.Rows[0]["Emp_No"].ToString(),dtEmp.Rows[0]["passwd"].ToString(),
                                                dtEmp.Rows[0]["Emp_Name"].ToString(),dtEmp.Rows[0]["Dept_Name"].ToString(),
                                                dtEmp.Rows[0]["Emp_No"].ToString(),dtEmp.Rows[0]["email"].ToString()};
                ado.dbNonQuery(insertStr, param);
                PageSwitch(Request.QueryString["tester"], sheet_categoryID, sheetID);
                return;
            }
            #endregion
            
            ScriptManager.RegisterStartupScript(this, this.GetType(), "js", "alert('account/password 輸入錯誤，請重新檢查');", true);
            txtAcc.Text = null;
            txtPw.Text = null;
            SetFocus(txtAcc);
        }
        private void PageSwitch(string tester, string sheet_categoryID, string sheetID)
        {
            try
            {
                string url = "";

                if (!string.IsNullOrEmpty(sheet_categoryID) && !string.IsNullOrEmpty(sheetID))
                {
                    url = string.Format(@"../eForm/Sheet_Values.aspx?sheet_categoryID={0}&sheetID={1}", sheet_categoryID, sheetID);
                }
                else if (!string.IsNullOrEmpty(tester))
                {
                    Common.AdoDbConn ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
                    string Am_sqlStr = string.Format(@"select acs_manageID from ACS_Manage where tester='{0}'", tester.ToUpper());
                    DataTable Am_dt = ado.loadDataTable(Am_sqlStr, null, "ACS_Manage");

                    string sqlStr = string.Format(@"Select sheet_categoryID from sheet_Category where tester like '%|{0}|%' And (isDelete is Null Or isDelete ='N')", Am_dt.Rows[0]["ACS_ManageID"].ToString());
                    DataTable dt = ado.loadDataTable(sqlStr, null, "Sheet_Category");
                    url = dt.Rows.Count > 0 ? string.Format(@"../eForm/Sheet_Values.aspx?sheet_categoryID={0}", dt.Rows[0]["sheet_categoryID"].ToString()) : url;
                }

                if (url == "")
                {
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('找不到此【{0}】機台對應資訊，請通知相關工程師。');", tester), true);
                    url = "../eForm/List.aspx";
                }
                Response.Redirect(string.Format(@"{0}", url));
            }
            catch (Exception ex) { }
        }



        #region demo
        private void ddl_bind()
        {
            Common.AdoDbConn ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            string str = @"Select describes,sheet_categoryID 
                           From sheet_category
                           Where (IsDelete IS null OR IsDelete != 'Y')
                           Order by Describes";
            DataTable dt = ado.loadDataTable(str, null, "sheet_category");

            ddl1.Items.Clear();
            foreach (DataRow dr in dt.Rows)
            {
                ddl1.Items.Add(new ListItem(dr["describes"].ToString(), dr["sheet_categoryID"].ToString()));
            }
            ddl1.Items.Insert(0, new ListItem("請選擇", "0"));

            ddl2.Items.Clear();
            ddl2.Items.Insert(0, new ListItem("請選擇", "0"));
        }
        protected void ddl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Common.AdoDbConn ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            string sqlStr = string.Format(@"select tester 
                                            From sheet_category 
                                            Where sheet_categoryid='{0}' And (IsDelete IS null OR IsDelete != 'Y')", ddl1.SelectedValue);
            DataTable dt = ado.loadDataTable(sqlStr, null, "sheet_category");
            string[] arrStr = dt.Rows[0]["tester"].ToString().Split('|');

            ddl2.Items.Clear();
            foreach (string str in arrStr)
            {
                if (!string.IsNullOrEmpty(str))
                {
                    string ddl2_sqlstr = string.Format(@"Select tester From ACS_Manage Where ACS_ManageID='{0}'", str);
                    DataTable ddl2_dt = ado.loadDataTable(ddl2_sqlstr, null, "ACS_Manage");
                    if(ddl2_dt.Rows.Count>0)
                        ddl2.Items.Add(new ListItem(ddl2_dt.Rows[0]["tester"].ToString()));
                }
            }
            ddl2.Items.Insert(0, new ListItem("請選擇", "0"));
        }
        protected void ddl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ddl1.SelectedValue.Equals("0")) return;
            Response.Redirect(string.Format(@"login.aspx?tester={0}", ddl2.SelectedValue));
        }
        #endregion
    }
}