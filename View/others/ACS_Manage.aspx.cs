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
using System.Text;
using System.Collections.Generic;

namespace EPM_Web.Alan
{
    public partial class ACS_Manage : System.Web.UI.Page
    {
        public static string Conn = System.Configuration.ConfigurationManager.ConnectionStrings["EPM"].ToString();
        public Common.AdoDbConn ado;
        public static DataTable dtGv;
        public static StringBuilder sb = new StringBuilder();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                Common.MyFunc func = new Common.MyFunc();
                func.checkLogin();
                func.checkRole(Page.Master);
                initial();
            }
        }

        private void initial()
        {
            gridviewBind();
        }
        private void gridviewBind()
        {
            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            DataTable dt = new DataTable();
            string str = string.Format(@"Select ACS_ManageID,tester,location,isSealing
                                         From   ACS_Manage
                                         Where  tester like '%{0}%'
                                                And location like '%{1}%'
                                         Order By ACS_ManageID desc", txtTester.Text.Trim().ToUpper(), txtLocation.Text.Trim().ToUpper());
            dt = ado.loadDataTable(str, null, "ACS_Manage");
            GridViewAM.DataSource = dt;
            GridViewAM.DataBind();
            dtGv = dt;
            showPage();
        }


        protected void btnSearch_Click(object sender, EventArgs e)
        {
            gridviewBind();
        }
        protected void btnNew_Click(object sender, EventArgs e)
        {
            GridViewAM.DataSource = null;
            GridViewAM.DataBind();
            DetailsView1.ChangeMode(DetailsViewMode.Insert);
        }
        protected void btnImport_Click(object sender, EventArgs e)
        {
            string acsConn = System.Configuration.ConfigurationManager.ConnectionStrings["ACS_ISA"].ToString();
            string sql = string.Format(@"Select tester,location
                                         From   VW_Address
                                         Where  location IS NOT NULL");
            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            Common.AdoDbConn acs_ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, acsConn); 
            DataTable acs_dt = acs_ado.loadDataTable(sql, null, "vw_Address");

            List<string> arrStr = new List<string>();
            foreach (DataRow dr in acs_dt.Rows)
            {
                string checksql = string.Format(@"Select  * 
                                                From    ACS_Manage
                                                Where   tester='{0}'", dr["tester"].ToString());
                DataTable checkdt = ado.loadDataTable(checksql, null, "ACS_Manage");
                if (checkdt.Rows.Count < 1)
                {
                    string transaction_Str = string.Format(@"Insert into ACS_Manage (acs_ManageID,Tester,Location,isSealing)
                                                          Values (ACS_manage_sequence.nextval,'{0}','{1}','N')", dr["Tester"].ToString(),
                                                                                                                 dr["Location"].ToString());
                    arrStr.Add(transaction_Str);
                }
            }
            string reStr = ado.SQL_transaction(arrStr, Conn);
            if (reStr.ToUpper().Contains("SUCCESS"))
                ScriptManager.RegisterStartupScript(this, this.GetType(), "js", "alert('匯入成功。');", true);
            else
                ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('{0}');", reStr), true);
        }

        protected void GridViewAM_Sorting(object sender, GridViewSortEventArgs e)
        {
            string sortExpression = e.SortExpression.ToString();
            string sortDirection = "ASC";

            if (sortExpression == this.GridViewAM.Attributes["SortExpression"])
                sortDirection = (this.GridViewAM.Attributes["SortDirection"].ToString() == sortDirection ? "DESC" : "ASC");

            this.GridViewAM.Attributes["SortExpression"] = sortExpression;
            this.GridViewAM.Attributes["SortDirection"] = sortDirection;

            if ((!string.IsNullOrEmpty(sortExpression)) && (!string.IsNullOrEmpty(sortDirection)))
            {
                dtGv.DefaultView.Sort = string.Format("{0} {1}", sortExpression, sortDirection);
            }
            GridViewAM.DataSource = dtGv;
            GridViewAM.DataBind();
            showPage();
        }
        protected void GridViewAM_RowEditing(object sender, GridViewEditEventArgs e)
        {
            GridViewAM.EditIndex = e.NewEditIndex;
            GridViewAM.DataSource = dtGv;
            GridViewAM.DataBind();
            showPage();
        }
        protected void GridViewAM_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
            GridViewAM.EditIndex = -1;
            GridViewAM.DataSource = dtGv;
            GridViewAM.DataBind();
            showPage();
        }
        protected void GridViewAM_RowUpdating(object sender, GridViewUpdateEventArgs e)
        {
            TextBox gv_txtTester = (TextBox)GridViewAM.Rows[e.RowIndex].Cells[0].FindControl("gv_txtTester");
            TextBox gv_txtLocation = (TextBox)GridViewAM.Rows[e.RowIndex].Cells[0].FindControl("gv_txtLocation");
            RadioButton gv_rdoIsSealing_Y = (RadioButton)GridViewAM.Rows[e.RowIndex].Cells[0].FindControl("gv_rdoIsSealing_Y");
            RadioButton gv_rdoIsSealing_N = (RadioButton)GridViewAM.Rows[e.RowIndex].Cells[0].FindControl("gv_rdoIsSealing_N");

            string pk = GridViewAM.DataKeys[e.RowIndex].Value.ToString();

            //check
            sb = new StringBuilder();
            bool check = true;

            if (string.IsNullOrEmpty(gv_txtTester.Text)) { sb.Append("【Tester】"); }
            if (string.IsNullOrEmpty(gv_txtLocation.Text)) { sb.Append("【Location】"); }
            if (gv_rdoIsSealing_Y.Checked == false && gv_rdoIsSealing_N.Checked == false) { sb.Append("【isSealing】"); }
            if (!string.IsNullOrEmpty(sb.ToString())) { check = false; sb.Insert(0, "以下欄位未填\\n\\n"); }

            if (check == true)
            {
                string oldStr = string.Format(@"Select tester,location From ACS_Manage Where ACS_Manageid='{0}'", pk);
                DataTable old_dt = ado.loadDataTable(oldStr, null, "ACS_MAnage");

                string TestStr = string.Format(@"Select * From ACS_Manage Where tester='{0}' And tester!='{1}'", gv_txtTester.Text.Trim().ToUpper(),old_dt.Rows[0]["tester"].ToString());
                DataTable tester_dt = ado.loadDataTable(TestStr, null, "ACS_Manage");
                if (tester_dt.Rows.Count > 0) { sb.Append("【Tester】"); }

                if (!string.IsNullOrEmpty(sb.ToString())) { check = false; sb.Insert(0, "以下欄位重複定義，請檢查\\n\\n"); }
            }

            if (check)
            {
                ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
                string str = string.Format(@"Update ACS_Manage 
                                             Set    tester = '{0}',
                                                    Location ='{1}',
                                                    isSealing ='{2}'
                                             Where  ACS_ManageID = '{3}'", gv_txtTester.Text.Trim().ToUpper(),
                                                                          gv_txtLocation.Text.Trim().ToUpper(),
                                                                          gv_rdoIsSealing_Y.Checked == true ? "Y" : "N",
                                                                          pk);
                //ado.dbNonQuery(str, null);
                string reStr = (string)ado.dbNonQuery(str, null);
                if (reStr.ToUpper().Contains("SUCCESS"))
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "js", "alert('該筆資料修改成功。');", true);
                else
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('{0}');", reStr), true); 

                GridViewAM.EditIndex = -1;
                gridviewBind();
            }
            else
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('{0}');", sb.ToString()), true);
            }
        }
        protected void GridViewAM_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            string sheet_categoryID = Request.QueryString["sheet_categoryID"];
            if (string.IsNullOrEmpty(sheet_categoryID)) { GridViewAM.Columns[6].Visible = false; }

            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                string pk = GridViewAM.DataKeys[e.Row.RowIndex].Value.ToString();

                Button gv_btnSelect = (Button)e.Row.FindControl("gv_btnSelect");
                ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
                                
                //=====================================================================
                //Selected
                //=====================================================================
                string str = string.Format(@"Select * 
                                      From   sheet_category 
                                      Where  tester like '%|{0}|%'
                                             And sheet_categoryID ='{1}' And (isDelete is Null Or isDelete ='N')", pk, sheet_categoryID);
                DataTable dtSelected = ado.loadDataTable(str, null, "sheet_Category");
                if (dtSelected.Rows.Count > 0) 
                { 
                    gv_btnSelect.Text = "Selected";
                    gv_btnSelect.BackColor = System.Drawing.Color.AliceBlue;
                    gv_btnSelect.Enabled = true; 
                }


                //=====================================================================
                //Occupied 
                //=====================================================================
                str = string.Format(@"Select * 
                                      From   sheet_category 
                                      Where  tester like '%|{0}|%'
                                             And sheet_categoryID != '{1}' And (isDelete is Null Or isDelete ='N')", pk, sheet_categoryID);
                DataTable dtOccupied = ado.loadDataTable(str, null, "sheet_Category");
                if (dtOccupied.Rows.Count > 0) 
                { 
                    gv_btnSelect.Text = "Occupied";
                    gv_btnSelect.BackColor = System.Drawing.Color.Red;
                    gv_btnSelect.Enabled = false; 
                }


                //=====================================================================
                //Available
                //=====================================================================
                str = string.Format(@"Select * 
                                      From   sheet_category 
                                      Where  tester like '%|{0}|%' And (isDelete is Null Or isDelete ='N')", pk);
                DataTable dtAvailable = ado.loadDataTable(str, null, "sheet_Category");
                if (dtAvailable.Rows.Count < 1) 
                {
                    gv_btnSelect.Text = "Available";
                    gv_btnSelect.BackColor = System.Drawing.Color.Green;
                    gv_btnSelect.Enabled = true; 
                }

                if (e.Row.RowState == DataControlRowState.Edit)
                {
                    RadioButton gv_rdoIsSealing_Y = (RadioButton)e.Row.FindControl("gv_rdoIsSealing_Y");
                    RadioButton gv_rdoIsSealing_N = (RadioButton)e.Row.FindControl("gv_rdoIsSealing_N");


                    //=====================================================================
                    //isSealing
                    //=====================================================================
                    str = string.Format(@"Select isSealing
                                          From   ACS_Manage 
                                          Where  ACS_ManageID ='{0}'", pk);
                    DataTable dtIsSealing = ado.loadDataTable(str, null, "ACS_Manage");
                    if (dtIsSealing.Rows.Count > 0)
                    {
                        if (dtIsSealing.Rows[0]["isSealing"].ToString().Equals("Y"))
                            gv_rdoIsSealing_Y.Checked = true;
                        else
                            gv_rdoIsSealing_N.Checked = true;
                    }
                }
            }
        }
        protected void GridViewAM_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            GridViewAM.PageIndex = e.NewPageIndex;
            GridViewAM.DataSource = dtGv;
            GridViewAM.DataBind();
            showPage();
        }
        protected void GridViewAM_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            string pk = e.CommandArgument.ToString();

            switch (e.CommandName)
            {
                case "myDelete":
                    string sqlStr = string.Format(@"Delete from ACS_Manage
                                                    Where ACS_ManageID='{0}'", pk);
                    //ado.dbNonQuery(sqlStr, null);
                    string reStr = (string)ado.dbNonQuery(sqlStr, null);
                    if (reStr.ToUpper().Contains("SUCCESS"))
                        ScriptManager.RegisterStartupScript(this, this.GetType(), "js", "alert('該筆資料刪除成功。');", true);
                    else
                        ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('{0}');", reStr), true); 

                    gridviewBind();
                    break;
                case "mySelect":
                    string sheet_categoryID = Request.QueryString["sheet_categoryID"];

                    GridViewRow row = ((Button)e.CommandSource).Parent.Parent as GridViewRow;
                    Button gv_btnSelect = (Button)row.FindControl("gv_btnSelect");

                    string strSelected = string.Format(@"Select tester From sheet_category Where sheet_categoryID='{0}'", sheet_categoryID);
                    DataTable dtSelected = ado.loadDataTable(strSelected, null, "sheet_Category");
                    string testerStr = dtSelected.Rows[0]["tester"].ToString();
                    string[] arrStr = testerStr.Split('|');

                    string Str = "";
                    switch (gv_btnSelect.Text)
                    {
                        //case "Occupied":
                        //    break;
                        case "Selected":
                            gv_btnSelect.Text = "Available";
                            gv_btnSelect.BackColor = System.Drawing.Color.Green;
                            foreach (string arr in arrStr)
                            {
                                if (!arr.Equals(pk) && !string.IsNullOrEmpty(arr)) { Str += (string.IsNullOrEmpty(Str) ? "|" + arr + "|" : arr + "|"); }
                            }
                            break;                        
                        case "Available":
                            gv_btnSelect.Text = "Selected";
                            gv_btnSelect.BackColor = System.Drawing.Color.AliceBlue;
                            Str += string.IsNullOrEmpty(testerStr) ? "|" + pk + "|" : testerStr + pk + "|";
                            break;
                    }
                    string updateStr = string.Format("Update sheet_category Set tester='{0}' where sheet_CategoryID='{1}'", Str, sheet_categoryID);
                    //ado.dbNonQuery(updateStr, null);
                    string up_reStr = (string)ado.dbNonQuery(updateStr, null);
                    if (up_reStr.ToUpper().Contains("SUCCESS"))
                        ScriptManager.RegisterStartupScript(this, this.GetType(), "js", "alert('該筆資料修改成功。');", true);
                    else
                        ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('{0}');", up_reStr), true); 
                    gridviewBind();
                    break;
            }
            
        }

        #region PagerTemplate
        protected void lbnFirst_Click(object sender, EventArgs e)
        {
            int num = 0;

            GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
            GridViewAM_PageIndexChanging(null, ea);
        }
        protected void lbnPrev_Click(object sender, EventArgs e)
        {
            int num = GridViewAM.PageIndex - 1;

            if (num >= 0)
            {
                GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
                GridViewAM_PageIndexChanging(null, ea);
            }
        }
        protected void lbnNext_Click(object sender, EventArgs e)
        {
            int num = GridViewAM.PageIndex + 1;

            if (num < GridViewAM.PageCount)
            {
                GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
                GridViewAM_PageIndexChanging(null, ea);
            }
        }
        protected void lbnLast_Click(object sender, EventArgs e)
        {
            int num = GridViewAM.PageCount - 1;

            GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
            GridViewAM_PageIndexChanging(null, ea);
        }
        private void showPage()
        {
            try
            {
                TextBox txtPage = (TextBox)GridViewAM.BottomPagerRow.FindControl("txtSizePage");
                Label lblCount = (Label)GridViewAM.BottomPagerRow.FindControl("lblTotalCount");
                Label lblPage = (Label)GridViewAM.BottomPagerRow.FindControl("lblPage");
                Label lblbTotal = (Label)GridViewAM.BottomPagerRow.FindControl("lblTotalPage");

                txtPage.Text = GridViewAM.PageSize.ToString();
                lblCount.Text = dtGv.Rows.Count.ToString();
                lblPage.Text = (GridViewAM.PageIndex + 1).ToString();
                lblbTotal.Text = GridViewAM.PageCount.ToString();

                GridViewAM.Columns[2].Visible = true;//IsNumeric
                GridViewAM.Columns[7].Visible = true;//MaxLimit
                GridViewAM.Columns[8].Visible = true;//MinLimit
            }
            catch (Exception ex) { }
        }
        // page change
        protected void btnGo_Click(object sender, EventArgs e)
        {
            try
            {
                string numPage = ((TextBox)GridViewAM.BottomPagerRow.FindControl("txtSizePage")).Text.ToString();
                if (!string.IsNullOrEmpty(numPage))
                {
                    GridViewAM.PageSize = Convert.ToInt32(numPage);
                }

                TextBox pageNum = ((TextBox)GridViewAM.BottomPagerRow.FindControl("inPageNum"));
                string goPage = pageNum.Text.ToString();
                if (!string.IsNullOrEmpty(goPage))
                {
                    int num = Convert.ToInt32(goPage) - 1;
                    if (num >= 0)
                    {
                        GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
                        GridViewAM_PageIndexChanging(null, ea);
                        ((TextBox)GridViewAM.BottomPagerRow.FindControl("inPageNum")).Text = null;
                    }
                }

                GridViewAM.DataSource = dtGv;
                GridViewAM.DataBind();
                showPage();
            }
            catch (Exception ex) { }
        }
        #endregion

        
        protected void DetailsView1_ItemInserting(object sender, DetailsViewInsertEventArgs e)
        {
            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            TextBox dv_txtTester = ((TextBox)DetailsView1.FindControl("dv_txtTester"));
            TextBox dv_txtLocation = ((TextBox)DetailsView1.FindControl("dv_txtLocation"));
            RadioButton dv_rdoIsSealing_Y = ((RadioButton)DetailsView1.FindControl("dv_rdoIsSealing_Y"));
            RadioButton dv_rdoIsSealing_N = ((RadioButton)DetailsView1.FindControl("dv_rdoIsSealing_N"));
            

            //Check
            sb = new StringBuilder();
            bool check = true;

            if (string.IsNullOrEmpty(dv_txtTester.Text)) { sb.Append("【Tester】"); }
            if (string.IsNullOrEmpty(dv_txtLocation.Text)) { sb.Append("【Location】"); }
            if (dv_rdoIsSealing_Y.Checked == false && dv_rdoIsSealing_N.Checked == false) { sb.Append("【isSealing】"); }
            if (!string.IsNullOrEmpty(sb.ToString())) { check = false; sb.Insert(0, "以下欄位未填\\n\\n"); }

            if (check == true)
            {
                string TestStr = string.Format(@"Select * From ACS_Manage Where tester='{0}'", dv_txtTester.Text.Trim().ToUpper());
                DataTable tester_dt = ado.loadDataTable(TestStr, null, "ACS_Manage");
                if (tester_dt.Rows.Count > 0) { sb.Append("【Tester】"); }
                if (!string.IsNullOrEmpty(sb.ToString())) { check = false; sb.Insert(0, "以下欄位重複定義，請檢查\\n\\n"); }
            }

            if (check)
            {
                string sqlStr = string.Format(@"Insert into ACS_Manage(ACS_ManageID,Tester,Location,isSealing)
                                                Values (ACS_manage_sequence.nextval,:Tester,:Location,:isSealing)");
                object[] para = new object[] { dv_txtTester.Text.Trim().ToUpper(), dv_txtLocation.Text.Trim().ToUpper(),
                                               dv_rdoIsSealing_Y.Checked==true?"Y":"N"};
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
                ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('{0}');", sb.ToString()), true);
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
        protected void DetailsView1_DataBound(object sender, EventArgs e)
        {
            if (DetailsView1.CurrentMode == DetailsViewMode.Insert)
            {
                //DropDownList ddl_DataType = (DropDownList)DetailsView1.FindControl("dv_ddlDataType");
            }
        }
        protected void DetailsView1_ModeChanging(object sender, DetailsViewModeEventArgs e)
        {

        }

        
}
}