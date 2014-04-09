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
    public partial class Sheet_Category : System.Web.UI.Page
    {
        public static string Conn = System.Configuration.ConfigurationManager.ConnectionStrings["EPM"].ToString();
        public Common.AdoDbConn ado;
        public static DataTable dtGv;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                Common.MyFunc func = new Common.MyFunc();
                func.checkLogin();
                func.checkRole(Page.Master);

                gridviewBind();


                #region trial-run db transport to production db

                //string newConn = System.Configuration.ConfigurationManager.ConnectionStrings["EPM_auto"].ToString();
                //List<string> arrStr = new List<string>();

//                int scID = 130;
//                string quStr = string.Format(@"select describes,tester,isdelete,docnumber from sheet_category where sheet_categoryId='{0}'", scID);
//                string quStr2 = string.Format(@"select itemid,datatypeid,sheet_categoryid,describes,method,spec,maxlimit,minlimit,value,
//                                         isurgent,isnumeric,iteminserttime,usersid,isdelete,isweekly,startWeek,frequency,ismonthly,startmonth,month_frequency 
//                                  from item where sheet_categoryId='{0}'", scID);
//                ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);

//                //sheet_category
//                DataTable dt = ado.loadDataTable(quStr, null, "sheet_category");
//                foreach (DataRow dr in dt.Rows)
//                {
//                    arrStr.Add(string.Format(@"insert into sheet_category(sheet_categoryID,describes,tester,isdelete,docnumber)
//                                               Values(sheet_category_sequence.nextval,'{0}','{1}','{2}','{3}')", dr["describes"].ToString(), dr["tester"].ToString(),
//                                                                                                                      dr["isdelete"].ToString(), dr["docnumber"].ToString()));
//                }

//                //item
//                DataTable dt2 = ado.loadDataTable(quStr2, null, "item");
//                foreach (DataRow dr in dt2.Rows)
//                {
//                    arrStr.Add(string.Format(@"insert into item(itemid,sheet_categoryid,describes,method,spec,maxlimit,minlimit,value,
//                                                                isurgent,isnumeric,iteminserttime,usersid,isdelete,
//                                                                isweekly,startWeek,frequency,ismonthly,startmonth,month_frequency,datatypeid)
//                                               Values(item_sequence.nextval,sheet_category_sequence.currval,'{0}','{1}','{2}','{3}','{4}','{5}'
//                                                      ,'{6}','{7}',SYSDATE,'{9}','{10}'
//                                                      ,'{11}','{12}','{13}','{14}','{15}','{16}','{17}')"
//                                                      , dr["describes"].ToString(), dr["method"].ToString(), dr["spec"].ToString(), dr["maxlimit"].ToString(), dr["minlimit"].ToString(), dr["value"].ToString()
//                                                      , dr["isurgent"].ToString(), dr["isnumeric"].ToString(), "", dr["usersid"].ToString(), dr["isdelete"].ToString()
//                                                      , dr["isweekly"].ToString(), dr["startWeek"].ToString(), dr["frequency"].ToString(), dr["ismonthly"].ToString(), dr["startmonth"].ToString(), dr["month_frequency"].ToString(), dr["datatypeid"].ToString()));
//                }

//                Common.AdoDbConn newado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, newConn);
//                if (newado.SQL_transaction(arrStr, newConn).ToUpper().Contains("SUCCESS")) { Console.WriteLine("success"); }
                //                else { Console.WriteLine("fail"); }
                #endregion
            }
        }

        private void gridviewBind()
        {
            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            DataTable dt = new DataTable();
//            string str = string.Format(
//                         @" Select sheet_categoryID,docNumber,describes,tester
//                            From  sheet_category
//                            WHERE (describes like '%{0}%' OR describes IS null) 
//                                  And (Tester like '%{1}%' OR Tester IS null)
//                                  And (docNumber like '%{2}%' OR docNumber IS null)
//                                  And (IsDelete IS null OR IsDelete != 'Y')
//                            Order by sheet_categoryID desc", txtDescribe.Text.Trim(), txtTester.Text.Trim().ToUpper(), txtDocNumber.Text.Trim());
            string str = string.Format(
                         @" Select sheet_categoryID,docNumber,describes,tester
                            From  sheet_category
                            WHERE (describes like '%{0}%' OR describes IS null) 
                                  And (docNumber like '%{1}%' OR docNumber IS null)
                                  And (IsDelete IS null OR IsDelete != 'Y')
                            Order by sheet_categoryID desc", txtDescribe.Text.Trim(),
                                                           txtDocNumber.Text.Trim());
            dt = ado.loadDataTable(str, null, "sheet_category");
            GridViewSC.DataSource = dt;
            GridViewSC.DataBind();
            dtGv = dt;
            showPage();
        }

        protected void btnSearch_Click(object sender, EventArgs e)
        {
            gridviewBind();
        }
        protected void GridViewSC_Sorting(object sender, GridViewSortEventArgs e)
        {
            string sortExpression = e.SortExpression.ToString();
            string sortDirection = "ASC";

            if (sortExpression == this.GridViewSC.Attributes["SortExpression"])
                sortDirection = (this.GridViewSC.Attributes["SortDirection"].ToString() == sortDirection ? "DESC" : "ASC");

            this.GridViewSC.Attributes["SortExpression"] = sortExpression;
            this.GridViewSC.Attributes["SortDirection"] = sortDirection;

            if ((!string.IsNullOrEmpty(sortExpression)) && (!string.IsNullOrEmpty(sortDirection)))
            {
                dtGv.DefaultView.Sort = string.Format("{0} {1}", sortExpression, sortDirection);
            }
            GridViewSC.DataSource = dtGv;
            GridViewSC.DataBind();
            showPage();
        }
        protected void GridViewSC_RowEditing(object sender, GridViewEditEventArgs e)
        {            
            GridViewSC.EditIndex = e.NewEditIndex;            
            GridViewSC.DataSource = dtGv;
            GridViewSC.DataBind();
            showPage();
        }
        protected void GridViewSC_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
            GridViewSC.EditIndex = -1;
            GridViewSC.DataSource = dtGv;
            GridViewSC.DataBind();
            showPage();
        }
        protected void GridViewSC_RowUpdating(object sender, GridViewUpdateEventArgs e)
        {
            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);

            string dk = GridViewSC.DataKeys[e.RowIndex].Value.ToString();
            string describe = ((TextBox)GridViewSC.Rows[e.RowIndex].Cells[0].FindControl("gv_txtDescribe")).Text.Trim();
            string docNumber = ((TextBox)GridViewSC.Rows[e.RowIndex].Cells[0].FindControl("gv_txtDocNumber")).Text.Trim();

            //Check
            string alertStr = string.Empty;
            bool check = true;
            if (string.IsNullOrEmpty(describe)) { check = false; alertStr = "【Describes】欄位未填"; }

            string checkStr = string.Format(@"  Select *
                                                From sheet_category 
                                                Where describes = '{0}' 
                                                      And sheet_categoryID !='{1}'
                                                      And (isDelete is Null or isDelete='N')", describe, dk);
            DataTable dt = ado.loadDataTable(checkStr, null, "sheet_category");
            if (dt.Rows.Count > 0)
            {
                check = false;
                alertStr = string.Format(@"該【{0}】表單名稱重複定義，請檢查。", dt.Rows[0]["describes"].ToString());
            }

            checkStr = string.Format(@"  Select * 
                                         From sheet_category 
                                         Where docNumber = '{0}' 
                                               And sheet_categoryID !='{1}'
                                               And (isDelete is Null or isDelete='N')", docNumber, dk);
            dt = ado.loadDataTable(checkStr, null, "sheet_category");
            if (dt.Rows.Count > 0)
            {
                check = false;
                alertStr = string.Format(@"該【{0}】文號重複定義，請檢查。", dt.Rows[0]["docNumber"].ToString());
            }


            if (check)
            {
                string str = string.Format(@"Update sheet_Category 
                                             Set describes='{0}',docNumber='{1}'
                                             Where sheet_categoryID='{2}'", describe, docNumber, dk);
                //ado.dbNonQuery(str, null);
                string reStr = (string)ado.dbNonQuery(str, null);
                if (reStr.ToUpper().Contains("SUCCESS"))
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "js", "alert('該筆資料修改成功。');", true);
                else
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('{0}');", reStr), true); 

                GridViewSC.EditIndex = -1;
                gridviewBind();
            }
            else
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('{0}');", alertStr), true);
            }
        }
        protected void GridViewSC_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            try
            {
                if (e.Row.RowType == DataControlRowType.DataRow)
                {
                    ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
                    TextBox gv_txtDescribe = (TextBox)e.Row.FindControl("gv_txtDescribe");
                    Label gv_Tester = (Label)e.Row.FindControl("gv_Tester");

                    string pk = GridViewSC.DataKeys[e.Row.RowIndex].Value.ToString();

                    string str = string.Format(@"select describes,tester from sheet_category WHERE sheet_categoryID='{0}'", pk);
                    DataTable dt = ado.loadDataTable(str, null, "sheet_category");
                    string[] arrStr = dt.Rows[0]["tester"].ToString().Split('|');

                    for (int i = 0; i < arrStr.Length; i++)
                    {
                        if (!string.IsNullOrEmpty(arrStr[i]))
                        {
                            string testStr = string.Format(@"Select tester From ACS_Manage Where ACS_ManageID='{0}'", arrStr[i]);
                            DataTable testDt = ado.loadDataTable(testStr, null, "ACS_Manage");
                            if (testDt.Rows.Count > 0)
                            {
                                gv_Tester.Text += string.Format(@"【{0}】", testDt.Rows[0]["tester"].ToString());
                                gv_Tester.Text += (i % 4 == 0) ? "<br/>" : "";

                                if ((e.Row.RowState & DataControlRowState.Edit) == DataControlRowState.Edit)
                                {
                                    gv_txtDescribe.TextMode = TextBoxMode.MultiLine;
                                    gv_txtDescribe.Rows = 3;
                                }
                            }
                        }
                    }

                }
            }
            catch (Exception ex) { }

        }
        protected void GridViewSC_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            GridViewSC.PageIndex = e.NewPageIndex;
            GridViewSC.DataSource = dtGv;
            GridViewSC.DataBind();
            showPage();
        }
        protected void GridViewSC_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            string pk = e.CommandArgument.ToString();
            List<string> arrStr = new List<string>();
            switch (e.CommandName)
            {
                case "Copy":
                    //===============================================================================
                    //Sheet_Category
                    //===============================================================================
                    string scStr = string.Format(@"Select describes,docNumber
                                               From sheet_Category
                                               Where sheet_categoryID='{0}'", pk);
                    DataTable dt_sc = ado.loadDataTable(scStr, null, "sheet_category");
                    arrStr.Add(string.Format(@"Insert into sheet_category (sheet_categoryID,describes)
                                               Values (sheet_category_sequence.nextval,'{0}')", dt_sc.Rows[0]["describes"].ToString() + "_(2)"));

                    //===============================================================================
                    //item
                    //===============================================================================
                    string itemStr = string.Format(@"Select datatypeID,sheet_categoryID,sheetid,describes,method,
                                                        spec,maxlimit,minlimit,value,isUrgent,IsNumeric,
                                                        isWeekly,StartWeek,Frequency,
                                                        isMonthly,StartMonth,Month_frequency
                                                 From item 
                                                 Where sheet_categoryID='{0}' 
                                                       And sheetID is null
                                                       And (isDelete ='N' or isDelete is Null)", pk);
                    DataTable dt_item = ado.loadDataTable(itemStr, null, "item");
                    foreach (DataRow dr in dt_item.Rows)
                    {
                        arrStr.Add(string.Format(@"Insert into item (itemid,datatypeid,sheet_categoryid,sheetid,
                                                                 describes,method,spec,maxlimit,minlimit,
                                                                 value,isUrgent,isNumeric, itemInsertTime,usersid,
                                                                 isWeekly,StartWeek,Frequency,
                                                                 isMonthly,StartMonth,Month_frequency)
                                               Values (item_sequence.nextval,'{0}',sheet_category_sequence.currval,'{1}',
                                                       '{2}','{3}','{4}','{5}','{6}',
                                                       '{7}','{8}','{9}',SYSDATE,'{10}',
                                                       '{11}','{12}','{13}',
                                                       '{14}','{15}','{16}')",
                                                   dr["DataTypeID"].ToString(), dr["sheetid"].ToString(),
                                                   dr["describes"].ToString(), dr["method"].ToString(), dr["spec"].ToString(), dr["maxlimit"].ToString(), dr["minlimit"].ToString(),
                                                   dr["value"].ToString(), dr["isUrgent"].ToString(), dr["isNumeric"].ToString(), Session["UsersID"],
                                                   dr["isWeekly"].ToString(), dr["StartWeek"].ToString(), dr["Frequency"].ToString(),
                                                   dr["isMonthly"].ToString(), dr["StartMonth"].ToString(), dr["Month_frequency"].ToString()));
                    }

                    string reStr = ado.SQL_transaction(arrStr, Conn);
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('{0}');", reStr), true);
                    gridviewBind();
                    break;

                case "myDelete":
                    string transaction_Str1 = string.Format(@"Update sheet_category 
                                                              SET    IsDelete='Y'
                                                              Where  sheet_CategoryID='{0}'", pk);

                    string transaction_Str2 = string.Format(@"Update item 
                                                              SET    IsDelete='Y'
                                                              Where  sheet_CategoryID='{0}'", pk);
                    arrStr.Add(transaction_Str1);
                    arrStr.Add(transaction_Str2);

                    string del_reStr = ado.SQL_transaction(arrStr, Conn);
                    if (del_reStr.ToUpper().Contains("SUCCESS"))
                        ScriptManager.RegisterStartupScript(this, this.GetType(), "js", "alert('該筆資料刪除成功。');", true);
                    else
                        ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('{0}');", del_reStr), true); 

                    gridviewBind();
                    break;
            }
            

        }
        #region PagerTemplate
        protected void lbnFirst_Click(object sender, EventArgs e)
        {
            int num = 0;

            GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
            GridViewSC_PageIndexChanging(null, ea);
        }
        protected void lbnPrev_Click(object sender, EventArgs e)
        {
            int num = GridViewSC.PageIndex - 1;

            if (num >= 0)
            {
                GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
                GridViewSC_PageIndexChanging(null, ea);
            }
        }
        protected void lbnNext_Click(object sender, EventArgs e)
        {
            int num = GridViewSC.PageIndex + 1;

            if (num < GridViewSC.PageCount)
            {
                GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
                GridViewSC_PageIndexChanging(null, ea);
            }
        }
        protected void lbnLast_Click(object sender, EventArgs e)
        {
            int num = GridViewSC.PageCount - 1;

            GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
            GridViewSC_PageIndexChanging(null, ea);
        }
        private void showPage()
        {
            try
            {
                TextBox txtPage = (TextBox)GridViewSC.BottomPagerRow.FindControl("txtSizePage");
                Label lblCount = (Label)GridViewSC.BottomPagerRow.FindControl("lblTotalCount");
                Label lblPage = (Label)GridViewSC.BottomPagerRow.FindControl("lblPage");
                Label lblbTotal = (Label)GridViewSC.BottomPagerRow.FindControl("lblTotalPage");

                txtPage.Text = GridViewSC.PageSize.ToString();
                lblCount.Text = dtGv.Rows.Count.ToString();
                lblPage.Text = (GridViewSC.PageIndex + 1).ToString();
                lblbTotal.Text = GridViewSC.PageCount.ToString();
            }
            catch (Exception ex) { }
        }
        // page change
        protected void btnGo_Click(object sender, EventArgs e)
        {
            try
            {
                string numPage = ((TextBox)GridViewSC.BottomPagerRow.FindControl("txtSizePage")).Text.ToString();
                if (!string.IsNullOrEmpty(numPage))
                {
                    GridViewSC.PageSize = Convert.ToInt32(numPage);
                }

                TextBox pageNum = ((TextBox)GridViewSC.BottomPagerRow.FindControl("inPageNum"));
                string goPage = pageNum.Text.ToString();
                if (!string.IsNullOrEmpty(goPage))
                {
                    int num = Convert.ToInt32(goPage) - 1;
                    if (num >= 0)
                    {
                        GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
                        GridViewSC_PageIndexChanging(null, ea);
                        ((TextBox)GridViewSC.BottomPagerRow.FindControl("inPageNum")).Text = null;
                    }
                }

                GridViewSC.DataSource = dtGv;
                GridViewSC.DataBind();
                showPage();
            }
            catch (Exception ex) { }
        }
        #endregion


        protected void btnNew_Click(object sender, EventArgs e)
        {
            GridViewSC.DataSource = null;
            GridViewSC.DataBind();
            DetailsView1.ChangeMode(DetailsViewMode.Insert);
        }
        protected void DetailsView1_ItemInserting(object sender, DetailsViewInsertEventArgs e)
        {
            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);

            TextBox dv_describes = ((TextBox)DetailsView1.FindControl("dv_describes"));
            TextBox dv_DocNumber = ((TextBox)DetailsView1.FindControl("dv_DocNumber"));

            //Check
            string alertStr = string.Empty;
            bool check = true;
            if (string.IsNullOrEmpty(dv_describes.Text)) { check = false; alertStr = "【Describes】欄位未填"; }

            string checkStr = string.Format(@"Select *
                                              From sheet_category 
                                              Where describes = '{0}' And (isDelete is Null or isDelete='N')", dv_describes.Text.Trim());
            
            DataTable dt = ado.loadDataTable(checkStr, null, "sheet_category");
            if (dt.Rows.Count > 0)
            {
                check = false;
                alertStr = string.Format(@"該【{0}】表單重複定義，請檢查。", dv_describes.Text.Trim());
            }

            checkStr = string.Format(@"Select *
                                       From sheet_category 
                                       Where docNumber = '{0}' And (isDelete is Null or isDelete='N')", dv_DocNumber.Text.Trim());

            dt = ado.loadDataTable(checkStr, null, "sheet_category");
            if (dt.Rows.Count > 0)
            {
                check = false;
                alertStr = string.Format(@"該【{0}】文號重複定義，請檢查。", dv_DocNumber.Text.Trim());
            }

            if (check)
            {
                string sqlStr = string.Format(@"Insert into sheet_category(sheet_categoryID,docNumber,describes)
                                                Values (sheet_category_sequence.nextval,:docNumber,:describes)");
                object[] para = new object[] { dv_DocNumber.Text.Trim(), dv_describes.Text.Trim() };
                //ado.dbNonQuery(sqlStr, para);
                string reStr = (string)ado.dbNonQuery(sqlStr, para);
                if (reStr.ToUpper().Contains("SUCCESS"))
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "js", "alert('該筆資料新增成功。');", true);
                else
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('{0}');", reStr), true); 

                DetailsView1.ChangeMode(DetailsViewMode.ReadOnly);
                DetailsView1.DataSource = null;
                DetailsView1.DataBind();

                gridviewBind();
            }
            else
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('{0}');", alertStr), true);
            }
            
        }        
        protected void DetailsView1_ItemCommand(object sender, DetailsViewCommandEventArgs e)
        {
            if (e.CommandName.Equals("cancel", StringComparison.CurrentCultureIgnoreCase))
            {
                DetailsView1.ChangeMode(DetailsViewMode.ReadOnly);
                DetailsView1.DataSource = null;
                DetailsView1.DataBind();

                gridviewBind();
            }
        }
        protected void DetailsView1_ModeChanging(object sender, DetailsViewModeEventArgs e)
        {

        }
        
}
}