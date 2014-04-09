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
    public partial class Item : System.Web.UI.Page
    {
        public static string Conn = System.Configuration.ConfigurationManager.ConnectionStrings["EPM"].ToString();
        public Common.AdoDbConn ado;
        public static DataTable dtGv;
        public static StringBuilder sb = new StringBuilder();

        public static string sheet_categoryID;
        public static string sheet_Category_describe;
        

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {                
                ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
               
                sheet_categoryID = Request.QueryString["sheet_categoryID"];
                string str = string.Format(@"select describes from sheet_category where sheet_categoryID='{0}'", sheet_categoryID);
                DataTable dt = ado.loadDataTable(str, null, "sheet_category");
                sheet_Category_describe = (dt.Rows.Count > 0) ? dt.Rows[0]["describes"].ToString() : "";

                Common.MyFunc func = new Common.MyFunc();
                func.checkLogin();
                func.checkRole(Page.Master);
                initial();                
            }
        }
        private void initial()
        {
            gridviewBind();

            //ddl
            string[] ddl_item = new string[] { "Y", "N" };
            foreach (string item in ddl_item)
            {
                ddlIsUrgent.Items.Add(new ListItem(item));
                ddlIsNumeric.Items.Add(new ListItem(item));
            }
            ddlIsUrgent.Items.Insert(0, new ListItem("請選擇"));
            ddlIsNumeric.Items.Insert(0, new ListItem("請選擇"));
        }
        private void gridviewBind()
        {
            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            DataTable dt = new DataTable();
            string str = string.Format(
                         @" Select itemID,item_desc,method,spec,maxlimit,minlimit,value,isurgent,isnumeric,iteminserttime,
                                   dataTypeID,datatype_desc,
                                   sheet_categoryID,sheetCategory_desc,
                                   sheetid,weekid,sheet_desc,
                                   usersid,roleid,account,password,name,dept,emp_no,mail,
                                   IsWeekly,StartWeek,Frequency,
                                   IsMonthly,StartMonth,Month_Frequency
                            From vw_item_column
                            Where sheet_categoryID='{0}' 
                                  And (item_desc like '%{1}%' or item_desc is null)
                                  And (method like '%{2}%' or method is null)
                                  And (spec like '%{3}%' or spec is null)
                                  And isUrgent like '%{4}%'
                                  And isNumeric like '%{5}%'
                                  And (item_IsDelete IS null OR item_IsDelete != 'Y')
                            Order by ItemID asc", sheet_categoryID,
                                            txtDescribe.Text.Trim(),
                                            txtMethod.Text.Trim(),
                                            txtSpec.Text.Trim(),
                                            (ddlIsUrgent.SelectedValue.Equals("請選擇")) ? "" : ddlIsUrgent.SelectedValue,
                                            (ddlIsNumeric.SelectedValue.Equals("請選擇")) ? "" : ddlIsNumeric.SelectedValue
                                            );
            dt = ado.loadDataTable(str, null, "vw_item_column");
            GridView1.DataSource = dt;
            GridView1.DataBind();
            dtGv = dt;
            showPage();

        }
        private DataTable getDatatype()
        {
            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            string str = @"select datatypeID,describes from datatype";
            DataTable dt = new DataTable();
            dt = ado.loadDataTable(str, null, "datatype");
            return dt;
        }

        
        protected void GridView1_Sorting(object sender, GridViewSortEventArgs e)
        {
            string sortExpression = e.SortExpression.ToString();
            string sortDirection = "ASC";

            if (sortExpression == this.GridView1.Attributes["SortExpression"])
                sortDirection = (this.GridView1.Attributes["SortDirection"].ToString() == sortDirection ? "DESC" : "ASC");

            this.GridView1.Attributes["SortExpression"] = sortExpression;
            this.GridView1.Attributes["SortDirection"] = sortDirection;

            if ((!string.IsNullOrEmpty(sortExpression)) && (!string.IsNullOrEmpty(sortDirection)))
            {
                dtGv.DefaultView.Sort = string.Format("{0} {1}", sortExpression, sortDirection);
            }
            GridView1.DataSource = dtGv;
            GridView1.DataBind();
            showPage();
        }
        protected void GridView1_RowEditing(object sender, GridViewEditEventArgs e)
        {
            GridView1.EditIndex = e.NewEditIndex;
            GridView1.DataSource = dtGv;
            GridView1.DataBind();
            showPage();
        }
        protected void GridView1_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
            GridView1.EditIndex = -1;
            GridView1.DataSource = dtGv;
            GridView1.DataBind();
            showPage();

            GridView1.Columns[2].Visible = true;//IsNumeric
            GridView1.Columns[7].Visible = true;//MaxLimit
            GridView1.Columns[8].Visible = true;//MinLimit
            GridView1.Columns[9].Visible = true;//In Week Base
            GridView1.Columns[10].Visible = true;//StartWeek
            GridView1.Columns[11].Visible = true;//Frequency
            GridView1.Columns[12].Visible = true;//In Month Base
            GridView1.Columns[13].Visible = true;//StartMonth
            GridView1.Columns[14].Visible = true;//Month_Frequency
        }
        protected void GridView1_RowUpdating(object sender, GridViewUpdateEventArgs e)
        {
            RadioButton gv_rdoIsUrgent_Y = (RadioButton)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_rdoIsUrgent_Y");
            RadioButton gv_rdoIsUrgent_N = (RadioButton)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_rdoIsUrgent_N");
            RadioButton gv_rdoIsNumeric_Y = (RadioButton)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_rdoIsNumeric_Y");
            RadioButton gv_rdoIsNumeric_N = (RadioButton)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_rdoIsNumeric_N");
            DropDownList gv_ddlDatatype = (DropDownList)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_ddlDatatype");
            RadioButton gv_rdoIsWeekly_Y = (RadioButton)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_rdoIsWeekly_Y");
            RadioButton gv_rdoIsWeekly_N = (RadioButton)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_rdoIsWeekly_N");
            RadioButton gv_rdoIsMonthly_Y = (RadioButton)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_rdoIsMonthly_Y");
            RadioButton gv_rdoIsMonthly_N = (RadioButton)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_rdoIsMonthly_N");
           

            string dk = GridView1.DataKeys[e.RowIndex].Value.ToString();
            string describes = ((TextBox)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_txtDescribe")).Text.Trim();
            string method = ((TextBox)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_txtMethod")).Text.Trim();
            string spec = ((TextBox)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_txtSpec")).Text.Trim();
            string maxlimit = ((TextBox)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_txtMaxlimit")).Text.Trim();
            string minlimit = ((TextBox)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_txtMinlimit")).Text.Trim();
            string IsUrgent = (gv_rdoIsUrgent_Y.Checked == true) ? "Y" : "N";
            string IsNumeric = (gv_rdoIsNumeric_Y.Checked == true) ? "Y" : "N";
            string datatypeID = gv_ddlDatatype.SelectedValue;
            string IsWeekly = (gv_rdoIsWeekly_Y.Checked == true) ? "Y" : "N";
            string StartWeek = ((TextBox)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_txtStartWeek")).Text.Trim();
            string Frequency = ((TextBox)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_txtFrequency")).Text.Trim();
            string IsMonthly = (gv_rdoIsMonthly_Y.Checked == true) ? "Y" : "N";
            string StartMonth = ((TextBox)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_txtStartMonth")).Text.Trim();
            string Month_Frequency = ((TextBox)GridView1.Rows[e.RowIndex].Cells[0].FindControl("gv_txtMonth_Frequency")).Text.Trim();
             
            #region Check
            Common.ValidatorHelper val = new Common.ValidatorHelper();
            sb = new StringBuilder();
            bool check = true;

            if (string.IsNullOrEmpty(describes)) { sb.Append("【Describes】"); }
            if (gv_rdoIsUrgent_Y.Checked == false && gv_rdoIsUrgent_N.Checked == false) { sb.Append("【IsUrgent】"); }
            if (datatypeID.Equals("請選擇")) { sb.Append("【DataType】"); }
            if ((gv_rdoIsWeekly_Y.Checked == false) && (gv_rdoIsWeekly_N.Checked == false)) { sb.Append("【IsWeekly】"); }
            if (gv_rdoIsWeekly_Y.Checked == true && string.IsNullOrEmpty(StartWeek)) { sb.Append("【StartWeek】"); }
            if (gv_rdoIsWeekly_Y.Checked == true && string.IsNullOrEmpty(Frequency)) { sb.Append("【Week Frequency】"); }
            if ((gv_rdoIsWeekly_Y.Checked == false) && (gv_rdoIsWeekly_N.Checked == false)) { sb.Append("【IsWeekly】"); }
            if (gv_rdoIsMonthly_Y.Checked == true && string.IsNullOrEmpty(StartMonth)) { sb.Append("【StartMonth】"); }
            if (gv_rdoIsMonthly_Y.Checked == true && string.IsNullOrEmpty(Month_Frequency)) { sb.Append("【Month Frequency】"); }
            if (!string.IsNullOrEmpty(sb.ToString())) { check = false; sb.Insert(0, "以下欄位未填\\n\\n"); }

            if (check)
            {
                if (gv_rdoIsWeekly_Y.Checked == true && gv_rdoIsMonthly_Y.Checked == true) { check = false; sb.Insert(0, "不能同時選取 【Is Week Base】 or 【Is Month Base】，僅能選擇一種。\\n\\n"); }
            }

            if (check)
            {
                if (gv_rdoIsNumeric_Y.Checked == true)
                {
                    if (!val.IsFloat(maxlimit) && (!string.IsNullOrEmpty(maxlimit))) { sb.Append("【MaxLimit】"); }
                    if (!val.IsFloat(minlimit) && (!string.IsNullOrEmpty(minlimit))) { sb.Append("【MinLimit】"); }
                }

                if (gv_rdoIsWeekly_Y.Checked == true)
                {
                    if (!val.IsFloat(StartWeek)) { sb.Append("【StartWeek】"); }
                    if (!val.IsFloat(Frequency)) { sb.Append("【Frequency】"); }
                }

                if (gv_rdoIsMonthly_Y.Checked == true)
                {
                    if (!val.IsFloat(StartMonth)) { sb.Append("【StartMonth】"); }
                    if (!val.IsFloat(Month_Frequency)) { sb.Append("【Month_Frequency】"); }
                }
                if (!string.IsNullOrEmpty(sb.ToString())) { check = false; sb.Insert(0, "以下欄位格式錯誤，請輸入數字\\n\\n"); }
            }

            #endregion

            if (check)
            {
                ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
                string str = string.Format(@"Update Item 
                                         Set 
                                             describes='{0}',method='{1}',spec='{2}',maxlimit='{3}',
                                             minlimit='{4}',IsUrgent='{5}',IsNumeric='{6}',datatypeID='{7}',
                                             IsWeekly='{8}',StartWeek='{9}',Frequency='{10}',
                                             IsMonthly='{11}',StartMonth='{12}',Month_Frequency='{13}'
                                         Where ItemID='{14}'", describes, method, spec, maxlimit,
                                                               minlimit, IsUrgent, IsNumeric, datatypeID,
                                                               IsWeekly, StartWeek, Frequency,
                                                               IsMonthly,StartMonth,Month_Frequency,
                                                               dk);
                //ado.dbNonQuery(str, null);
                string reStr = (string)ado.dbNonQuery(str, null);
                if (reStr.ToUpper().Contains("SUCCESS"))
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "js", "alert('該筆資料修改成功。');", true);
                else
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('{0}');", reStr), true); 


                GridView1.EditIndex = -1;
                gridviewBind();
            }
            else
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('{0}');", sb.ToString()), true);
            }

            GridView1.Columns[2].Visible = true;//IsNumeric
            GridView1.Columns[7].Visible = true;//MaxLimit
            GridView1.Columns[8].Visible = true;//MinLimit
            GridView1.Columns[9].Visible = true;//In Week Base
            GridView1.Columns[10].Visible = true;//StartWeek
            GridView1.Columns[11].Visible = true;//Week Frequency
            GridView1.Columns[12].Visible = true;//In Month Base
            GridView1.Columns[13].Visible = true;//StartMonth
            GridView1.Columns[14].Visible = true;//Month Frequency
        }
        protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow
                    && (e.Row.RowState & DataControlRowState.Edit) == DataControlRowState.Edit)
            {
                TextBox gv_txtDescribe = (TextBox)e.Row.FindControl("gv_txtDescribe");
                TextBox gv_txtMethod = (TextBox)e.Row.FindControl("gv_txtMethod");
                TextBox gv_txtSpec = (TextBox)e.Row.FindControl("gv_txtSpec");
                TextBox gv_txtMaxlimit = (TextBox)e.Row.FindControl("gv_txtMaxlimit");
                TextBox gv_txtMinlimit = (TextBox)e.Row.FindControl("gv_txtMinlimit");
                RadioButton gv_rdoIsUrgent_Y = (RadioButton)e.Row.FindControl("gv_rdoIsUrgent_Y");
                RadioButton gv_rdoIsUrgent_N = (RadioButton)e.Row.FindControl("gv_rdoIsUrgent_N");
                RadioButton gv_rdoIsNumeric_Y = (RadioButton)e.Row.FindControl("gv_rdoIsNumeric_Y");
                RadioButton gv_rdoIsNumeric_N = (RadioButton)e.Row.FindControl("gv_rdoIsNumeric_N");
                DropDownList gv_ddlDatatype = (DropDownList)e.Row.FindControl("gv_ddlDatatype");
                RadioButton gv_rdoIsWeekly_Y = (RadioButton)e.Row.FindControl("gv_rdoIsWeekly_Y");
                RadioButton gv_rdoIsWeekly_N = (RadioButton)e.Row.FindControl("gv_rdoIsWeekly_N");
                RadioButton gv_rdoIsMonthly_Y = (RadioButton)e.Row.FindControl("gv_rdoIsMonthly_Y");
                RadioButton gv_rdoIsMonthly_N = (RadioButton)e.Row.FindControl("gv_rdoIsMonthly_N"); 

                ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
                string dk = GridView1.DataKeys[e.Row.RowIndex].Value.ToString();
                string str = string.Format(@" Select itemID,item_desc,method,spec,maxlimit,minlimit,value,isurgent,isnumeric,iteminserttime,
                                                     dataTypeID,datatype_desc,
                                                     sheet_categoryID,sheetCategory_desc,
                                                     sheetid,weekid,sheet_desc,
                                                     usersid,roleid,account,password,name,dept,emp_no,mail,
                                                     IsWeekly,Startweek,Frequency,
                                                     IsMonthly,StartMonth,Month_Frequency
                                              From vw_item_column
                                              Where sheet_categoryID='{0}' 
                                                    And itemID='{1}'
                                              Order by ItemID", sheet_categoryID,
                                                                dk);
                DataTable dt = ado.loadDataTable(str, null, "vw_item_column");

                gv_txtDescribe.Text = dt.Rows[0]["item_desc"].ToString();
                gv_txtMethod.Text = dt.Rows[0]["method"].ToString();
                gv_txtSpec.Text = dt.Rows[0]["spec"].ToString();
                gv_txtMaxlimit.Text = dt.Rows[0]["maxlimit"].ToString();
                gv_txtMinlimit.Text = dt.Rows[0]["minlimit"].ToString();

                DataTable dt_datatype = getDatatype();
                foreach (DataRow dr in dt_datatype.Rows)
                {
                    gv_ddlDatatype.Items.Add(new ListItem(dr["describes"].ToString(), dr["datatypeID"].ToString()));
                }
                gv_ddlDatatype.SelectedValue = dt.Rows[0]["dataTypeID"].ToString();
                if (!gv_ddlDatatype.SelectedItem.Text.Contains("TextBox"))
                {
                    GridView1.Columns[2].Visible = false;//IsNumeric
                    GridView1.Columns[7].Visible = false;//MaxLimit
                    GridView1.Columns[8].Visible = false;//MinLimit
                }

                if (dt.Rows[0]["isurgent"].ToString().Equals("Y")) { gv_rdoIsUrgent_Y.Checked = true; }
                else { gv_rdoIsUrgent_N.Checked = true; }

                if (dt.Rows[0]["isnumeric"].ToString().Equals("Y")) 
                { gv_rdoIsNumeric_Y.Checked = true; }
                else
                {
                    gv_rdoIsNumeric_N.Checked = true; 
                    GridView1.Columns[7].Visible = false;
                    GridView1.Columns[8].Visible = false;
                }

                if (dt.Rows[0]["IsWeekly"].ToString().Equals("Y"))
                {
                    gv_rdoIsWeekly_Y.Checked = true;
                    //GridView1.Columns[12].Visible = false;//Is Month Base
                    GridView1.Columns[13].Visible = false;//StartMonth
                    GridView1.Columns[14].Visible = false;//Month Frequency
                }
                else
                {
                    gv_rdoIsWeekly_N.Checked = true;
                    //GridView1.Columns[12].Visible = true;//Is Month Base
                    GridView1.Columns[10].Visible = false;//StartWeek
                    GridView1.Columns[11].Visible = false;//Week Frequency
                }

                if (dt.Rows[0]["IsMonthly"].ToString().Equals("Y"))
                {
                    gv_rdoIsMonthly_Y.Checked = true;
                    //GridView1.Columns[9].Visible = false;//Is Week Base
                    GridView1.Columns[10].Visible = false;//StartWeek
                    GridView1.Columns[11].Visible = false;//Week Frequency
                }
                else
                {
                    gv_rdoIsMonthly_N.Checked = true;
                    //GridView1.Columns[9].Visible = true;//Is Week Base
                    GridView1.Columns[13].Visible = false;//StartMonth
                    GridView1.Columns[14].Visible = false;//Month Frequency
                }
            }
        }
        protected void GridView1_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            GridView1.PageIndex = e.NewPageIndex;
            GridView1.DataSource = dtGv;
            GridView1.DataBind();
            showPage();
        }
        protected void GridView1_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            string pk = e.CommandArgument.ToString();

            switch (e.CommandName)
            {
                case "myDelete":
                    string sqlStr = string.Format(@"Update item 
                                                    SET    IsDelete='Y' 
                                                    Where  itemID='{0}'", pk);
                    //ado.dbNonQuery(sqlStr, null);
                    string reStr = (string)ado.dbNonQuery(sqlStr, null);
                    if (reStr.ToUpper().Contains("SUCCESS"))
                        ScriptManager.RegisterStartupScript(this, this.GetType(), "js", "alert('該筆資料刪除成功。');", true);
                    else
                        ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('{0}');", reStr), true); 

                    gridviewBind();
                    break;
            }
        }
        protected void gv_ddlDatatype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (((DropDownList)sender).SelectedItem.Text.Contains("TextBox"))
            {
                GridView1.Columns[2].Visible = true;//IsNumeric
                GridView1.Columns[7].Visible = true;//MaxLimit
                GridView1.Columns[8].Visible = true;//MinLimit
            }
            else
            {
                GridView1.Columns[2].Visible = false;
                GridView1.Columns[7].Visible = false;
                GridView1.Columns[8].Visible = false;
            }
        }
        protected void gv_rdoIsNumeric_Y_CheckedChanged(object sender, EventArgs e)
        {
            GridView1.Columns[7].Visible = true;//MaxLimit
            GridView1.Columns[8].Visible = true;//MinLimit
        }
        protected void gv_rdoIsNumeric_N_CheckedChanged(object sender, EventArgs e)
        {
            GridView1.Columns[7].Visible = false;
            GridView1.Columns[8].Visible = false;
        }
        protected void gv_rdoIsWeekly_Y_CheckedChanged(object sender, EventArgs e)
        {
            GridView1.Columns[10].Visible = true;//StartWeek
            GridView1.Columns[11].Visible = true;//Week Frequency
        }
        protected void gv_rdoIsWeekly_N_CheckedChanged(object sender, EventArgs e)
        {
            GridView1.Columns[10].Visible = false;//StartWeek
            GridView1.Columns[11].Visible = false;//Week Frequency
            GridView1.Columns[12].Visible = true;//Is Month Base
        }
        protected void gv_rdoIsMonthly_Y_CheckedChanged(object sender, EventArgs e)
        {
            GridView1.Columns[13].Visible = true;//StartMonth
            GridView1.Columns[14].Visible = true;//Month Frequency
        }
        protected void gv_rdoIsMonthly_N_CheckedChanged(object sender, EventArgs e)
        {
            GridView1.Columns[9].Visible = true;//Is Week Base
            GridView1.Columns[13].Visible = false;//StartMonth
            GridView1.Columns[14].Visible = false;//Month Frequency
        }

        #region PagerTemplate
        protected void lbnFirst_Click(object sender, EventArgs e)
        {
            int num = 0;

            GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
            GridView1_PageIndexChanging(null, ea);
        }
        protected void lbnPrev_Click(object sender, EventArgs e)
        {
            int num = GridView1.PageIndex - 1;

            if (num >= 0)
            {
                GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
                GridView1_PageIndexChanging(null, ea);
            }
        }
        protected void lbnNext_Click(object sender, EventArgs e)
        {
            int num = GridView1.PageIndex + 1;

            if (num < GridView1.PageCount)
            {
                GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
                GridView1_PageIndexChanging(null, ea);
            }
        }
        protected void lbnLast_Click(object sender, EventArgs e)
        {
            int num = GridView1.PageCount - 1;

            GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
            GridView1_PageIndexChanging(null, ea);
        }
        private void showPage()
        {
            try
            {
                TextBox txtPage = (TextBox)GridView1.BottomPagerRow.FindControl("txtSizePage");
                Label lblCount = (Label)GridView1.BottomPagerRow.FindControl("lblTotalCount");
                Label lblPage = (Label)GridView1.BottomPagerRow.FindControl("lblPage");
                Label lblbTotal = (Label)GridView1.BottomPagerRow.FindControl("lblTotalPage");

                txtPage.Text = GridView1.PageSize.ToString();
                lblCount.Text = dtGv.Rows.Count.ToString();
                lblPage.Text = (GridView1.PageIndex + 1).ToString();
                lblbTotal.Text = GridView1.PageCount.ToString();

                //GridView1.Columns[2].Visible = true;//IsNumeric
                //GridView1.Columns[7].Visible = true;//MaxLimit
                //GridView1.Columns[8].Visible = true;//MinLimit
            }
            catch (Exception ex) { }
        }
        // page change
        protected void btnGo_Click(object sender, EventArgs e)
        {
            try
            {
                string numPage = ((TextBox)GridView1.BottomPagerRow.FindControl("txtSizePage")).Text.ToString();
                if (!string.IsNullOrEmpty(numPage))
                {
                    GridView1.PageSize = Convert.ToInt32(numPage);
                }

                TextBox pageNum = ((TextBox)GridView1.BottomPagerRow.FindControl("inPageNum"));
                string goPage = pageNum.Text.ToString();
                if (!string.IsNullOrEmpty(goPage))
                {
                    int num = Convert.ToInt32(goPage) - 1;
                    if (num >= 0)
                    {
                        GridViewPageEventArgs ea = new GridViewPageEventArgs(num);
                        GridView1_PageIndexChanging(null, ea);
                        ((TextBox)GridView1.BottomPagerRow.FindControl("inPageNum")).Text = null;
                    }
                }

                GridView1.DataSource = dtGv;
                GridView1.DataBind();
                showPage();
            }
            catch (Exception ex) { }
        }
        #endregion


        protected void btnSearch_Click(object sender, EventArgs e)
        {
            gridviewBind();
        }
        protected void btnNew_Click(object sender, EventArgs e)
        {
            GridView1.DataSource = null;
            GridView1.DataBind();
            DetailsView1.ChangeMode(DetailsViewMode.Insert);
        }

        protected void DetailsView1_ItemInserting(object sender, DetailsViewInsertEventArgs e)
        {
            TextBox dv_describes = ((TextBox)DetailsView1.FindControl("dv_describes"));
            DropDownList ddl_DataType = (DropDownList)DetailsView1.FindControl("dv_ddlDataType");
            TextBox dv_Method = ((TextBox)DetailsView1.FindControl("dv_Method"));
            TextBox dv_Spec = ((TextBox)DetailsView1.FindControl("dv_Spec"));
            TextBox dv_MaxLimit = ((TextBox)DetailsView1.FindControl("dv_MaxLimit"));
            TextBox dv_MinLimit = ((TextBox)DetailsView1.FindControl("dv_MinLimit"));
            RadioButton dv_RdoIsUrgent_Y = ((RadioButton)DetailsView1.FindControl("dv_RdoIsUrgent_Y"));
            RadioButton dv_RdoIsUrgent_N = ((RadioButton)DetailsView1.FindControl("dv_RdoIsUrgent_N"));
            RadioButton dv_RdoIsNumeric_Y = ((RadioButton)DetailsView1.FindControl("dv_RdoIsNumeric_Y"));
            RadioButton dv_RdoIsNumeric_N = ((RadioButton)DetailsView1.FindControl("dv_RdoIsNumeric_N"));
            RadioButton dv_rdoIsWeekly_Y = ((RadioButton)DetailsView1.FindControl("dv_rdoIsWeekly_Y"));
            RadioButton dv_rdoIsWeekly_N = ((RadioButton)DetailsView1.FindControl("dv_rdoIsWeekly_N"));
            TextBox dv_StartWeek = ((TextBox)DetailsView1.FindControl("dv_StartWeek"));
            TextBox dv_Frequency = ((TextBox)DetailsView1.FindControl("dv_Frequency"));
            RadioButton dv_rdoIsMonthly_Y = ((RadioButton)DetailsView1.FindControl("dv_rdoIsMonthly_Y"));
            RadioButton dv_rdoIsMonthly_N = ((RadioButton)DetailsView1.FindControl("dv_rdoIsMonthly_N"));
            TextBox dv_StartMonth = ((TextBox)DetailsView1.FindControl("dv_StartMonth"));
            TextBox dv_Month_Frequency = ((TextBox)DetailsView1.FindControl("dv_Month_Frequency"));

            #region Check
            Common.ValidatorHelper val = new Common.ValidatorHelper();
            sb = new StringBuilder();
            bool check = true;

            if (ddl_DataType.SelectedItem.Text.Contains("請選擇")) { sb.Append("【DataType】"); }
            if (ddl_DataType.SelectedItem.Text.Contains("TextBox") && (dv_RdoIsNumeric_Y.Checked == false) && (dv_RdoIsNumeric_N.Checked == false)) { sb.Append("【IsNumeric】"); }
            if ((dv_RdoIsUrgent_Y.Checked == false) && (dv_RdoIsUrgent_N.Checked == false)) { sb.Append("【IsUrgent】"); }
            if (string.IsNullOrEmpty(dv_describes.Text)) { sb.Append("【Describes】"); }
            if ((dv_rdoIsWeekly_Y.Checked == false) && (dv_rdoIsWeekly_N.Checked == false)) { sb.Append("【IsWeekly】"); }
            if (dv_rdoIsWeekly_Y.Checked == true && string.IsNullOrEmpty(dv_StartWeek.Text)) { sb.Append("【StartWeek】"); }
            if (dv_rdoIsWeekly_Y.Checked == true && string.IsNullOrEmpty(dv_Frequency.Text)) { sb.Append("【Frequency】"); }
            if ((dv_rdoIsMonthly_Y.Checked == false) && (dv_rdoIsMonthly_N.Checked == false)) { sb.Append("【IsMonthly】"); }
            if (dv_rdoIsMonthly_Y.Checked == true && string.IsNullOrEmpty(dv_StartMonth.Text)) { sb.Append("【StartMonth】"); }
            if (dv_rdoIsMonthly_Y.Checked == true && string.IsNullOrEmpty(dv_Month_Frequency.Text)) { sb.Append("【Month_Frequency】"); }
            if (!string.IsNullOrEmpty(sb.ToString())) { check = false; sb.Insert(0, "以下欄位未填\\n\\n"); }

            if (check)
            {
                if (dv_rdoIsMonthly_Y.Checked == true && dv_rdoIsMonthly_Y.Checked == true) { check = false; sb.Insert(0, "不能同時選取 【Is Week Base】 or 【Is Month Base】，僅能選擇一種。\\n\\n"); }
            }

            if (check)
            {
                if (dv_RdoIsNumeric_Y.Checked == true)
                {
                    if (!val.IsFloat(dv_MaxLimit.Text.Trim()) && (!string.IsNullOrEmpty(dv_MaxLimit.Text.Trim()))) { sb.Append("【MaxLimit】"); }
                    if (!val.IsFloat(dv_MinLimit.Text.Trim()) && (!string.IsNullOrEmpty(dv_MinLimit.Text.Trim()))) { sb.Append("【MinLimit】"); }
                }

                if (dv_rdoIsWeekly_Y.Checked == true)
                {
                    if (!val.IsFloat(dv_StartWeek.Text.Trim())) { sb.Append("【StartWeek】"); }
                    if (!val.IsFloat(dv_Frequency.Text.Trim())) { sb.Append("【Frequency】"); }
                }

                if (dv_rdoIsMonthly_Y.Checked == true)
                {
                    if (!val.IsFloat(dv_StartMonth.Text.Trim())) { sb.Append("【StartMonth】"); }
                    if (!val.IsFloat(dv_Month_Frequency.Text.Trim())) { sb.Append("【Month_Frequency】"); }
                }
                if (!string.IsNullOrEmpty(sb.ToString())) { check = false; sb.Insert(0, "以下欄位格式錯誤，請輸入數字\\n\\n"); }
            }
            #endregion

            if (check)
            {
                ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
                string sqlStr = string.Format(@"Insert into item(itemID,DataTypeID,Sheet_CategoryID,describes,
                                                                Method,Spec,MaxLimit,MinLimit,
                                                                IsUrgent,IsNumeric,
                                                                itemInsertTime,UsersID,
                                                                IsWeekly,StartWeek,Frequency,
                                                                IsMonthly,StartMonth,Month_Frequency)
                                            Values (item_sequence.nextval,:DataTypeID,:Sheet_CategoryID,:describes,
                                                    :Method,:Spec,:MaxLimit,:MinLimit,
                                                    :IsUrgent,:IsNumeric,
                                                    SYSDATE,:UsersID,
                                                    :IsWeekly,:StartWeek,:Frequency,
                                                    :IsMonthly,:StartMonth,:Month_Frequency)");
                object[] para = new object[] { ddl_DataType.SelectedValue,sheet_categoryID, dv_describes.Text.Trim(),
                                            dv_Method.Text.Trim(),dv_Spec.Text.Trim(),dv_MaxLimit.Text.Trim(),dv_MinLimit.Text.Trim(),
                                            (dv_RdoIsUrgent_Y.Checked)?"Y":"N",(dv_RdoIsNumeric_Y.Checked)?"Y":"N",
                                            Session["UsersID"],
                                            (dv_rdoIsWeekly_Y.Checked)?"Y":"N",dv_StartWeek.Text.Trim(),dv_Frequency.Text.Trim(),
                                            (dv_rdoIsMonthly_Y.Checked)?"Y":"N",dv_StartMonth.Text.Trim(),dv_Month_Frequency.Text.Trim()};
                //ado.dbNonQuery(sqlStr, para);

                string reStr = (string)ado.dbNonQuery(sqlStr, para);
                if (reStr.ToUpper().Contains("SUCCESS"))
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "js", "alert('該筆資料新增成功。');", true);
                else
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('新增失敗，該動作出現Exception。');"), true);

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
                DropDownList ddl_DataType = (DropDownList)DetailsView1.FindControl("dv_ddlDataType");
                ddl_DataType.Items.Clear();
                DataTable dt_dataType = getDatatype();
                foreach (DataRow dr in dt_dataType.Rows)
                {
                    ddl_DataType.Items.Add(new ListItem(dr["describes"].ToString(), dr["datatypeID"].ToString()));
                }
                ddl_DataType.Items.Insert(0, new ListItem("請選擇"));
            }
        }
        protected void DetailsView1_ModeChanging(object sender, DetailsViewModeEventArgs e)
        {

        }

        protected void dv_ddlDataType_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddl_DataType = (DropDownList)DetailsView1.FindControl("dv_ddlDataType");
            TextBox dv_MaxLimit = ((TextBox)DetailsView1.FindControl("dv_MaxLimit"));
            TextBox dv_MinLimit = ((TextBox)DetailsView1.FindControl("dv_MinLimit"));
            RadioButton dv_RdoIsNumeric_Y = ((RadioButton)DetailsView1.FindControl("dv_RdoIsNumeric_Y"));
            RadioButton dv_RdoIsNumeric_N = ((RadioButton)DetailsView1.FindControl("dv_RdoIsNumeric_N"));

            if (ddl_DataType.SelectedItem.Text.Contains("TextBox"))
            {                
                DetailsView1.Rows[1].Visible = true;//IsNumeric
                DetailsView1.Rows[6].Visible = true;//MaxLimit
                dv_MaxLimit.Text = null;
                DetailsView1.Rows[7].Visible = true;//MinLimit   
                dv_MinLimit.Text = null;
            }
            else
            {
                dv_RdoIsNumeric_Y.Checked = false;
                dv_RdoIsNumeric_N.Checked = false;
                DetailsView1.Rows[1].Visible = false;
                DetailsView1.Rows[6].Visible = false;
                dv_MaxLimit.Text = null;
                DetailsView1.Rows[7].Visible = false;
                dv_MinLimit.Text = null;
            }
        }
        protected void dv_RdoIsNumeric_N_CheckedChanged(object sender, EventArgs e)
        {
            TextBox dv_MaxLimit = ((TextBox)DetailsView1.FindControl("dv_MaxLimit"));
            TextBox dv_MinLimit = ((TextBox)DetailsView1.FindControl("dv_MinLimit"));

            DetailsView1.Rows[6].Visible = false;
            dv_MaxLimit.Text = null;
            DetailsView1.Rows[7].Visible = false;
            dv_MinLimit.Text = null;
        }
        protected void dv_RdoIsNumeric_Y_CheckedChanged(object sender, EventArgs e)
        {
            DetailsView1.Rows[6].Visible = true;
            DetailsView1.Rows[7].Visible = true;
        }
        protected void dv_rdoIsWeekly_Y_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton dv_rdoIsMonthly_N = ((RadioButton)DetailsView1.FindControl("dv_rdoIsMonthly_N"));
            TextBox dv_StartMonth = ((TextBox)DetailsView1.FindControl("dv_StartMonth"));
            TextBox dv_Month_Frequency = ((TextBox)DetailsView1.FindControl("dv_Month_Frequency"));
            TextBox dv_StartWeek = ((TextBox)DetailsView1.FindControl("dv_StartWeek"));
            TextBox dv_Frequency = ((TextBox)DetailsView1.FindControl("dv_Frequency"));

            dv_rdoIsMonthly_N.Checked = true;
            dv_StartWeek.Enabled = true;
            dv_Frequency.Enabled = true;
            dv_StartMonth.Enabled = false;
            dv_Month_Frequency.Enabled = false;
            dv_StartMonth.Text = null;
            dv_Month_Frequency.Text = null;

        }
        protected void dv_rdoIsWeekly_N_CheckedChanged(object sender, EventArgs e)
        {
            TextBox dv_StartWeek = ((TextBox)DetailsView1.FindControl("dv_StartWeek"));
            TextBox dv_Frequency = ((TextBox)DetailsView1.FindControl("dv_Frequency"));

            dv_StartWeek.Enabled = false;
            dv_Frequency.Enabled = false;
            dv_StartWeek.Text = null;
            dv_Frequency.Text = null;
        }
        protected void dv_rdoIsMonthly_Y_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton dv_rdoIsWeekly_N = ((RadioButton)DetailsView1.FindControl("dv_rdoIsWeekly_N"));
            TextBox dv_StartWeek = ((TextBox)DetailsView1.FindControl("dv_StartWeek"));
            TextBox dv_Frequency = ((TextBox)DetailsView1.FindControl("dv_Frequency"));
            TextBox dv_StartMonth = ((TextBox)DetailsView1.FindControl("dv_StartMonth"));
            TextBox dv_Month_Frequency = ((TextBox)DetailsView1.FindControl("dv_Month_Frequency"));

            dv_rdoIsWeekly_N.Checked = true;
            dv_StartMonth.Enabled = true;
            dv_Month_Frequency.Enabled = true;
            dv_StartWeek.Enabled = false;
            dv_Frequency.Enabled = false;
            dv_StartWeek.Text = null;
            dv_Frequency.Text = null;
        }
        protected void dv_rdoIsMonthly_N_CheckedChanged(object sender, EventArgs e)
        {
            TextBox dv_StartMonth = ((TextBox)DetailsView1.FindControl("dv_StartMonth"));
            TextBox dv_Month_Frequency = ((TextBox)DetailsView1.FindControl("dv_Month_Frequency"));

            dv_StartMonth.Enabled = false;
            dv_Month_Frequency.Enabled = false;
            dv_StartMonth.Text = null;
            dv_Month_Frequency.Text = null;
        }

        protected void btnImportCsv_Click(object sender, EventArgs e)
        {
            try
            {
                if (FileUpload1.HasFile)
                {
                    string filename = FileUpload1.FileName;

                    string saveDir = System.Configuration.ConfigurationManager.AppSettings["UploadDir"].ToString();
                    string appPath = Request.PhysicalApplicationPath;
                    string savePath = appPath + saveDir;

                    string saveResult = savePath + filename;

                    FileUpload1.SaveAs(saveResult);
                    //ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('上傳成功\n 檔案名稱:{0}');", filename ), true);


                    Common.MyFunc func = new Common.MyFunc();
                    DataTable dtCsv = func.readCsvTxt(saveResult);

                    List<string> arrStr = new List<string>();
                    for (int i = 0; i < dtCsv.Rows.Count; i++)
                    {
                        if (i > 1)
                        {
                            DataRow dr = dtCsv.Rows[i];
                            string sql = string.Format(@"insert into item(itemID,datatypeID,describes,method,spec,sheet_categoryID,isUrgent,isNumeric,isWeekly,isMonthly,usersID,ItemInsertTime)
                                                 Values(item_sequence.nextval,'{0}','{1}','{2}','{3}','{4}','Y','N','Y','Y','{5}',SYSDATE)",
                                                            dr["datatype"].ToString(), dr["describes"].ToString(), dr["method"].ToString(), dr["spec"].ToString(), sheet_categoryID, Session["UsersID"]);
                            arrStr.Add(sql);
                        }
                    }

                    ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
                    string reStr = ado.SQL_transaction(arrStr, Conn);
                    if (reStr.ToUpper().Contains("SUCCESS"))
                        ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('csv匯入成功\n 檔案名稱:{0}');", filename), true);
                    else
                        ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('csv匯入失敗:{0}');", reStr), true);

                    gridviewBind();
                }
                else
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('請選擇檔案再上傳');"), true);
            }
            catch (Exception ce) { ScriptManager.RegisterStartupScript(this, this.GetType(), "js", string.Format(@"alert('Import fail..Please Download User Manual and refer to page29~33.');"), true); }
        }

        
}
}