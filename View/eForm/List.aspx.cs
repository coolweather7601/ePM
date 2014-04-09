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
using System.Data.OracleClient;
using System.Text;
using System.Collections.Generic;

namespace EPM_Web.Alan
{
    public partial class List : System.Web.UI.Page
    {
        public static string Conn = System.Configuration.ConfigurationManager.ConnectionStrings["EPM"].ToString();
        public Common.AdoDbConn ado;        
        public static DataTable dtGv;
        public static string title_sheet_category;
        public static string sheetID, sheet_categoryID;


        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                Common.MyFunc func = new Common.MyFunc();
                func.checkLogin();
                func.checkRole(Page.Master);

                //ddl databind
                ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
                string sc_str = @"Select sheet_CategoryID,describes 
                                  From   sheet_Category
                                  Where  (IsDelete IS null OR IsDelete != 'Y')
                                  Order by describes";
                DataTable dt_sc = ado.loadDataTable(sc_str, null, "sheet_Category");                
                foreach (DataRow dr in dt_sc.Rows)
                {
                    ddl_sheetCategory.Items.Add(new ListItem(dr["describes"].ToString(), dr["sheet_CategoryID"].ToString()));
                }

                ddl_sheetCategory.Items.Insert(0, new ListItem("請選擇", ""));
                btnSearch_Click(null, null);
            }
        }
        private void deleteData(string deleteID)
        {
            try
            {
                //==================================================================================================
                //log(insert/edit/delete)
                //==================================================================================================
                string log_str = @"Insert into Insert_delete_Log (insert_delete_logID,sheetID,UsersID,Log_action,Log_Time)
                                   Values (log_sequence.nextval,:SheetID,:UsersID,'DELETE',:log_Time)";
                object[] log_para = new object[] { deleteID, Session["UsersID"].ToString(), DateTime.Now };
                ado.dbNonQuery(log_str, log_para);

                //==================================================================================================
                //delete Data
                //==================================================================================================
                List<string> arrStr= new List<string>();
                string transaction_Str1 = "", transaction_Str2 = "";
                transaction_Str1 = string.Format(@"Update item set IsDelete='Y' where sheetId='{0}'", deleteID);
                transaction_Str2 = string.Format(@"Update sheet set IsDelete='Y' where sheetId='{0}'", deleteID);
                arrStr.Add(transaction_Str1);
                arrStr.Add(transaction_Str2);
                ado.SQL_transaction(arrStr, Conn);
                ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('已成功刪除該筆資料');", true);
                btnSearch_Click(null, null);
            }
            catch (Exception ex) { }
        }
      
        protected void GridViewList_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            sheetID = e.CommandArgument.ToString();

            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            string sqlstr = string.Format(@"select sheet_categoryID from sheet where sheetID='{0}'", sheetID);
            DataTable dt = ado.loadDataTable(sqlstr, null, "sheet");

            if (e.CommandName.Equals("view")) { ScriptManager.RegisterStartupScript(this, this.GetType(), "", string.Format(@"window.open('sheet_Values.aspx?sheetID={0}&sheet_categoryID={1}','_blank','width=1200,height=600,scrollbars=yes');", sheetID, dt.Rows[0]["sheet_categoryID"].ToString()), true); }
            if (e.CommandName.Equals("del")) { deleteData(e.CommandArgument.ToString()); }            
            if (e.CommandName.Equals("log"))
            {
                string str = string.Format(@"../others/UpdateLog.aspx?sheetID={0}", sheetID);
                ScriptManager.RegisterStartupScript(this, this.GetType(), "", string.Format(@"window.open('{0}','_blank','width=700,height=600,scrollbars=yes');",str), true);
            }
        }
        protected void GridViewList_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            GridViewList.PageIndex = e.NewPageIndex;
            GridViewList.DataSource = dtGv;
            GridViewList.DataBind();
            showPage();
        }
        protected void GridViewList_RowEditing(object sender, GridViewEditEventArgs e)
        {
            GridViewList.EditIndex = e.NewEditIndex;
            GridViewList.DataSource = dtGv;
            GridViewList.DataBind();
            showPage();
        }
        protected void GridViewList_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
            GridViewList.EditIndex = -1;
            GridViewList.DataSource = dtGv;
            GridViewList.DataBind();
            showPage();
        }
        protected void GridViewList_RowUpdating(object sender, GridViewUpdateEventArgs e)
        {
            string dk = GridViewList.DataKeys[e.RowIndex].Value.ToString(); //sheetID
            string remark = ((TextBox)GridViewList.Rows[e.RowIndex].Cells[0].FindControl("gv_txtMemo")).Text;

            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            string str = string.Format(@"Update sheet
                                         Set    Memo=:Memo
                                         Where  sheetID='{0}'", dk);
            object[] para = new object[] { remark };
            //ado.dbNonQuery(str, para);
            string reStr = (string)ado.dbNonQuery(str, para);
            if (reStr.ToUpper().Contains("SUCCESS"))
                ScriptManager.RegisterStartupScript(this, this.GetType(), "js", "alert('該筆資料修改成功。');", true);
            else
                ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('{0}');", reStr), true); 


            GridViewList.EditIndex = -1;
            btnSearch_Click(null, null);
        }
        protected void GridViewList_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);

            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                //=========================================================================
                //Edit mode
                //=========================================================================
                if ((e.Row.RowState & DataControlRowState.Edit) == DataControlRowState.Edit)
                {

                    //string str = "select RemarkID, Describes from Remark";
                    //DataTable dt = ado.loadDataTable(str, null, "Remark");

                    //DropDownList ddl = (DropDownList)e.Row.FindControl("gv_ddlRemark");
                    //ddl.Items.Clear();
                    //foreach (DataRow dr in dt.Rows)
                    //{
                    //    ddl.Items.Add(new ListItem(dr["Describes"].ToString(), dr["RemarkID"].ToString()));
                    //}

                    //string dk = GridViewList.DataKeys[e.Row.RowIndex].Value.ToString();
                    //str = string.Format(@"select RemarkID from sheet WHERE sheetID='{0}'", dk);
                    //dt = ado.loadDataTable(str, null, "sheet");
                    //ddl.SelectedValue = dt.Rows[0]["RemarkID"].ToString();
                }
            }
        }
        #region PagerTemplate
        protected void lbnFirst_Click(object sender, EventArgs e)
        {
            int num = 0;

            GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
            GridViewList_PageIndexChanging(null, ea);
        }
        protected void lbnPrev_Click(object sender, EventArgs e)
        {
            int num = GridViewList.PageIndex - 1;

            if (num >= 0)
            {
                GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
                GridViewList_PageIndexChanging(null, ea);
            }
        }
        protected void lbnNext_Click(object sender, EventArgs e)
        {
            int num = GridViewList.PageIndex + 1;

            if (num < GridViewList.PageCount)
            {
                GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
                GridViewList_PageIndexChanging(null, ea);
            }
        }
        protected void lbnLast_Click(object sender, EventArgs e)
        {
            int num = GridViewList.PageCount - 1;

            GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
            GridViewList_PageIndexChanging(null, ea);
        }
        private void showPage()
        {
            try
            {
                TextBox txtPage = (TextBox)GridViewList.BottomPagerRow.FindControl("txtSizePage");
                Label lblCount = (Label)GridViewList.BottomPagerRow.FindControl("lblTotalCount");
                Label lblPage = (Label)GridViewList.BottomPagerRow.FindControl("lblPage");
                Label lblbTotal = (Label)GridViewList.BottomPagerRow.FindControl("lblTotalPage");

                txtPage.Text = GridViewList.PageSize.ToString();
                lblCount.Text = dtGv.Rows.Count.ToString();
                lblPage.Text = (GridViewList.PageIndex + 1).ToString();
                lblbTotal.Text = GridViewList.PageCount.ToString();
            }
            catch (Exception ex) { }
        }
        // page change
        protected void btnGo_Click(object sender, EventArgs e)
        {
            try
            {
                string numPage = ((TextBox)GridViewList.BottomPagerRow.FindControl("txtSizePage")).Text.ToString();
                if (!string.IsNullOrEmpty(numPage))
                {
                    GridViewList.PageSize = Convert.ToInt32(numPage);
                }

                TextBox pageNum = ((TextBox)GridViewList.BottomPagerRow.FindControl("inPageNum"));
                string goPage = pageNum.Text.ToString();
                if (!string.IsNullOrEmpty(goPage))
                {
                    int num = Convert.ToInt32(goPage) - 1;
                    if (num >= 0)
                    {
                        GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
                        GridViewList_PageIndexChanging(null, ea);
                        ((TextBox)GridViewList.BottomPagerRow.FindControl("inPageNum")).Text = null;
                    }
                }

                GridViewList.DataSource = dtGv;
                GridViewList.DataBind();
                showPage();
            }
            catch (Exception ex) { }
        }
        #endregion


        protected void ddl_sheetCategory_SelectedIndexChanged(object sender, EventArgs e)
        {
            sheet_categoryID = ddl_sheetCategory.SelectedValue;
            btnSearch_Click(null, null);
        }
        protected void btnSearch_Click(object sender, EventArgs e)
        {
            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            ContentPlaceHolder placeHold = (ContentPlaceHolder)Master.FindControl("ContentPlaceHolder2");
            DataTable dt = new DataTable();

            string sqlStr = string.Format(@"Select vw.sheetID,vw.sheet_categoryID,vw.sheetCategory_desc,vw.remarkID,vw.remark_desc,
                                                   vw.weekID,vw.machine,vw.location,vw.memo,vw.insertTime,ad.machine_owner as owner1,ad2.machine_owner as owner2
                                            From vw_sheet vw
                                            left join address@ACS_OPE1.ACS ad on vw.machine=ad.tester
                                            left join address@ACS_OPE2.ACS ad2 on vw.machine=ad2.tester
                                            where vw.sheet_categoryID = '{0}' {1}
                                                  And (vw.IsDelete IS null OR vw.IsDelete != 'Y')
                                                  And vw.Machine like '%{2}%'
                                                  And vw.Location like '%{3}%' {4} {5}
                                                  And {6}
                                            Order by weekid desc ,sheetid desc", sheet_categoryID,
                                                                               (string.IsNullOrEmpty(sheet_categoryID)) ? "OR sheet_categoryID like '%' " : "",
                                                                               txtTester.Text.Trim().ToUpper(),
                                                                               txtLocation.Text.Trim().ToUpper(),
                                                                               (!string.IsNullOrEmpty(txtStartdate.Text)) ? "And weekID >= '" + txtStartdate.Text + "'" : "",
                                                                               (!string.IsNullOrEmpty(txtEnddate.Text)) ? "And weekID <= '" + txtEnddate.Text + "'" : "",
                                                                               string.IsNullOrEmpty(txtOwner.Text) ? "(ad.machine_owner is null or ad2.machine_owner is null)" 
                                                                                                                   : string.Format(@"(ad.machine_owner like '%{0}%' or ad2.machine_owner like '%{0}%')",txtOwner.Text.ToUpper())
                                                                                                                   );
            

            dt = ado.loadDataTable(sqlStr, null, "vw_sheet");

            GridViewList.DataSource = dt;
            GridViewList.DataBind();
            dtGv = dt;
            showPage();
        }
}
}