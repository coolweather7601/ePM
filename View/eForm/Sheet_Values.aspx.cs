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
    public partial class Sheet_Values : System.Web.UI.Page
    {
        public static string Conn = System.Configuration.ConfigurationManager.ConnectionStrings["EPM"].ToString();
        public Common.AdoDbConn ado;
        public static string title_sheetCategory_desc;
        public static StringBuilder sb_NullOrEmpty = new StringBuilder();
        public static StringBuilder sb_Validity = new StringBuilder();
        public static StringBuilder sb_Urgent = new StringBuilder();
        public static StringBuilder sb_Abnormal = new StringBuilder();

        public static string sheet_categoryID, sheetID;
        //if(string.isNull(sheetID)){Insert_Mode} else { Update_Mode}

        protected void Page_Init(object sender, EventArgs e)
        {
            sheet_categoryID = Request.QueryString["sheet_categoryID"];
            sheetID = Request.QueryString["sheetID"];

            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);

            string str = string.Format(@"SELECT itemID,item_desc,method,spec,Maxlimit,Minlimit,value,datatypeID,
                                                datatype_desc,sheetCategory_desc,machine,location,weekid,
                                                IsWeekly,StartWeek,Frequency,
                                                isMonthly,StartMonth,Month_Frequency
                                         FROM vw_item
                                         WHERE sheet_categoryID='{0}' {1}
                                         Order by itemID asc",
                                         sheet_categoryID,
                                         string.IsNullOrEmpty(sheetID) ? "And SheetID is null And (sheet_IsDelete IS null OR sheet_IsDelete != 'Y') And (item_IsDelete IS null OR item_IsDelete != 'Y')"
                                                                       : string.Format(@"And SheetID='{0}'", sheetID));
            DataTable dt = ado.loadDataTable(str, null, "vw_item");

            #region controls & htmltable

            HtmlTable html_tb = new HtmlTable();
            HtmlTableRow html_thead = new HtmlTableRow();
            HtmlTableRow html_tfoot = new HtmlTableRow();
            HtmlTableRow html_tr = new HtmlTableRow();
            HtmlTableCell html_cell = new HtmlTableCell();
            PH.Controls.Add(html_tb);

            //==================================================================================================
            //table
            //==================================================================================================
            html_tb.Border = 1;
            html_tb.ID = "contentTable";
            html_tb.Attributes.Add("class", "table table-bordered");
            html_tb.Attributes.Add("style", "width:70%;");


            //==================================================================================================
            //thead
            //==================================================================================================
            string[] thead = new string[] { "項次", "檢查點", "檢查方法", "基準", "Value" };
            string[] thead_width = new string[] { "3%", "20%", "25%", "27%", "25%" };
            
            for (int i = 0; i < thead.Length; i++)
            {
                html_cell = new HtmlTableCell();
                html_cell.InnerText = thead[i];
                html_cell.Attributes.Add("style", string.Format(@"background-color: #aaaaaa; text-align: center;width:{0};", thead_width[i]));
                html_thead.Controls.Add(html_cell); 
            }
            html_tb.Controls.Add(html_thead);


            //==================================================================================================
            //tbody
            //==================================================================================================
            int item_count = 1;
            foreach (DataRow dr in dt.Rows)
            {
                //檢查IsWeekly Or IsMonthly Item(Insert Mode) 
                //(該週週別 or 月份) - (起始週別 or 月份) 除以 Frequency
                //若餘數不為0，則表示該item不為當週 or 當月檢查項目
                bool WeeklyOrMonthly = true;                
                if (string.IsNullOrEmpty(sheetID) && string.IsNullOrEmpty(Request.QueryString["preview"])) 
                {
                    Common.MyFunc func = new Common.MyFunc();

                    #region Is Week Base OR Is Month Base

                    if (dr["IsWeekly"].ToString().Equals("Y") && dr["IsMonthly"].ToString().Equals("N"))//weekly item
                    {
                        int nowWeekID = Convert.ToInt32(func.getWeekCode(DateTime.Now));
                        //int nowWeekID = string.IsNullOrEmpty(Request.QueryString["TestDate"]) ? Convert.ToInt32(func.getWeekCode(DateTime.Now)) :
                        //                                                                        Convert.ToInt32(func.getWeekCode(DateTime.Parse(Request.QueryString["TestDate"])));
                        int startWeekID = Convert.ToInt32(dr["StartWeek"].ToString());
                        int frequency = Convert.ToInt32(dr["frequency"].ToString());
                        if ((nowWeekID - startWeekID) % frequency != 0) { WeeklyOrMonthly = false; }//if no remainder, show items
                    }
                    else if (dr["IsMonthly"].ToString().Equals("Y") && dr["IsWeekly"].ToString().Equals("N"))//monthly item
                    {
                        string getMonthID = func.getMonthCode(DateTime.Now);
                        //string getMonthID = string.IsNullOrEmpty(Request.QueryString["TestDate"]) ? func.getMonthCode(DateTime.Now)
                        //                                                                          : func.getMonthCode(DateTime.Parse(Request.QueryString["TestDate"]));

                        //if no remainder, show items(check first week in this month?)
                        if (string.IsNullOrEmpty(getMonthID))
                        {
                            WeeklyOrMonthly = false;
                        }
                        else
                        {
                            int nowMonthID = Convert.ToInt32(getMonthID);
                            int startMonthID = Convert.ToInt32(dr["StartMonth"].ToString());
                            int Month_frequency = Convert.ToInt32(dr["Month_frequency"].ToString());

                            //if no remainder, show items(check first week in this month?)
                            if ((nowMonthID - startMonthID) % Month_frequency != 0) { WeeklyOrMonthly = false; }
                        }
                    }
                    #endregion
                }

                if (WeeklyOrMonthly)
                {
                    txtOperator.Text = Session["account"] != null ? Session["account"].ToString() : "";
                    txtMachine.Text = dr["machine"].ToString();
                    txtLocation.Text = dr["location"].ToString();
                    txtWeek.Text = dr["weekid"].ToString();

                    html_tr = new HtmlTableRow();

                    html_cell = new HtmlTableCell();
                    html_cell.InnerText = item_count.ToString();
                    html_tr.Controls.Add(html_cell);

                    html_cell = new HtmlTableCell();
                    html_cell.InnerText = dr["Item_desc"].ToString();
                    html_cell.Attributes.Add("class", "tdbgt");
                    html_tr.Controls.Add(html_cell);

                    html_cell = new HtmlTableCell();
                    html_cell.InnerText = dr["method"].ToString();
                    html_tr.Controls.Add(html_cell);

                    html_cell = new HtmlTableCell();
                    html_cell.InnerText = dr["spec"].ToString();
                    html_tr.Controls.Add(html_cell);

                    switch (dr["datatype_desc"].ToString())
                    {
                        case "TextBox":
                            html_cell = new HtmlTableCell();
                            TextBox txt = new TextBox();
                            txt.ID = dr["ItemID"].ToString();
                            txt.Attributes.Add("onkeydown", "preventTextEnterEvent();");
                            html_cell.Controls.Add(txt);
                            txt.Text = dr["Value"].ToString();
                            break;
                        case "RadioButton":
                            html_cell = new HtmlTableCell();

                            RadioButton rdo_y = new RadioButton();
                            rdo_y.ID = string.Format(@"rdo_y_{0}", dr["ItemID"].ToString());
                            rdo_y.Text = "Yes";
                            rdo_y.GroupName = dr["ItemID"].ToString();
                            html_cell.Controls.Add(rdo_y);

                            RadioButton rdo_n = new RadioButton();
                            rdo_n.ID = string.Format(@"rdo_n_{0}", dr["ItemID"].ToString());
                            rdo_n.Text = "Normal";
                            rdo_n.GroupName = dr["ItemID"].ToString();
                            html_cell.Controls.Add(rdo_n);

                            RadioButton rdo_a = new RadioButton();
                            rdo_a.ID = string.Format(@"rdo_a_{0}", dr["ItemID"].ToString());
                            rdo_a.Text = "Abnormal";
                            rdo_a.GroupName = dr["ItemID"].ToString();
                            html_cell.Controls.Add(rdo_a);

                            if (!string.IsNullOrEmpty(dr["Value"].ToString()))
                            {
                                if (dr["Value"].ToString().Equals("Y")) rdo_y.Checked = true;
                                else if (dr["Value"].ToString().Equals("N")) rdo_n.Checked = true;
                                else if (dr["Value"].ToString().Equals("A")) rdo_a.Checked = true;
                            }
                            break;
                        //case "DropDownList":
                        //    html_cell = new HtmlTableCell();
                        //    DropDownList ddl = new DropDownList();
                        //    ddl.ID = dr["ItemID"].ToString();

                        //    DataTable lst_dt = getListItem(sheet_categoryID);
                        //    foreach (DataRow lst_dr in lst_dt.Rows) { ddl.Items.Add(new ListItem(lst_dr["describes"].ToString())); }
                        //    ddl.Items.Insert(0, new ListItem("請選擇", "00000"));
                        //    html_cell.Controls.Add(ddl);
                        //    ddl.SelectedValue = dr["Value"].ToString();
                        //    break;
                        default:
                            break;
                    }
                    html_tr.Controls.Add(html_cell);
                    html_tb.Controls.Add(html_tr);
                    item_count++;
                }
            }

            //==================================================================================================
            //tfoot (preview/edit)
            //==================================================================================================
            
            html_cell = new HtmlTableCell();
            html_cell.Attributes.Add("colspan", "5");
            if (string.IsNullOrEmpty(Request.QueryString["preview"]))
            {
                Button btnSubmit = new Button();
                btnSubmit.ID = "btnSubmit";
                btnSubmit.Text = "Submit";
                btnSubmit.Attributes.Add("CssClass", "btn btn-primary");
                btnSubmit.Attributes.Add("onclick", "return confirm('確定送出表單?');");
                btnSubmit.Click += new EventHandler(btnSubmit_Click);
                html_cell.Controls.Add(btnSubmit);
                html_tfoot.Controls.Add(html_cell);

                Button btnReset = new Button();
                btnReset.ID = "btnReset";
                btnReset.Text = "Reset";
                btnReset.Attributes.Add("CssClass", "btn btn-inverse");
                btnReset.Attributes.Add("onclick", "return confirm('確定把資料清空?');");
                btnReset.Click += new EventHandler(btnReset_Click);
                html_cell.Controls.Add(btnReset);
                html_tfoot.Controls.Add(html_cell);
            }
            Button btnMail = new Button();
            btnMail.ID = "btnMail";
            btnMail.Text = "Mail";
            btnMail.Attributes.Add("CssClass", "btn btn-warning");
            btnMail.Attributes.Add("onclick", string.Format(@"window.open('../others/MailSet.aspx?sheet_categoryID={0}','_blank','width=1100,height=700,scrollbars=yes');", sheet_categoryID));
            html_cell.Controls.Add(btnMail);
            html_tfoot.Controls.Add(html_cell);
            html_tb.Controls.Add(html_tfoot);
           
            #endregion

            title_sheetCategory_desc = (dt.Rows.Count > 0) ? dt.Rows[0]["sheetCategory_desc"].ToString() : "";
        }
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                Common.MyFunc func = new Common.MyFunc();
                func.checkLogin();
                func.checkRole(Page.Master);

                if (string.IsNullOrEmpty(sheetID)) { func.getAcsMachineData(Page.Master); }
                else
                {
                    ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
                    string str = string.Format(@"SELECT machine,location,weekid,account 
                                                 FROM   vw_sheet
                                                 WHERE  sheetID=:sheetID");
                    object[] para = new object[] { sheetID };
                    DataTable dt = ado.loadDataTable(str, para, "vw_sheet");
                    foreach (DataRow dr in dt.Rows)
                    {
                        txtMachine.Text = dr["machine"].ToString();
                        txtLocation.Text = dr["location"].ToString();
                        txtOperator.Text = dr["account"].ToString();
                        txtWeek.Text = dr["weekid"].ToString();
                    }                    
                }
            }
        }
        

        protected void btnReset_Click(object sender, EventArgs e)
        {
            ContentPlaceHolder PlaceHolder2 = (ContentPlaceHolder)Page.Master.FindControl("ContentPlaceHolder2");
           
            PlaceHolder panel = (PlaceHolder)PlaceHolder2.FindControl("PH");
            HtmlTable html_tb = (HtmlTable)panel.FindControl("contentTable");

            foreach (object ctr in html_tb.Controls)
            {
                foreach (object obj in ((HtmlTableRow)ctr).Controls)//HtmlTableRow
                {
                    foreach (object objs in ((HtmlTableCell)obj).Controls)//HtmlTableCell
                    {
                        switch (objs.GetType().Name.ToString())
                        {
                            case "TextBox":
                                ((TextBox)objs).Text = "";
                                break;
                            case "RadioButton":
                                ((RadioButton)objs).Checked = false;
                                break;
                            //case "DropDownList":
                            //    if (((DropDownList)objs).Items.Count > 0) { ((DropDownList)objs).SelectedIndex = 0; }
                            //    break;
                        }                        
                    }
                }                
            }

        }
        protected void btnSubmit_Click(object sender, EventArgs e)
        {
            try
            {
                Common.MyFunc func = new Common.MyFunc();
                func.checkLogin();

                sb_NullOrEmpty = new StringBuilder();
                sb_Validity = new StringBuilder();
                sb_Urgent = new StringBuilder();
                sb_Abnormal = new StringBuilder();

                ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
                string str = string.Format(@"SELECT itemID,item_desc,method,spec,Maxlimit,Minlimit,value,datatypeID,
                                                datatype_desc,sheetCategory_desc,sheet_categoryID,sheetID,iteminserttime,
                                                usersid,weekid,isUrgent,isNumeric,
                                                isWeekly,StartWeek,Frequency,
                                                isMonthly,StartMonth,Month_frequency
                                         FROM vw_item
                                         WHERE sheet_categoryID='{0}' {1}",
                                             sheet_categoryID,
                                             string.IsNullOrEmpty(sheetID) ? "And SheetID is null And (sheet_IsDelete IS null OR sheet_IsDelete != 'Y') And (item_IsDelete IS null OR item_IsDelete != 'Y')"
                                                                           : string.Format(@"And SheetID='{0}'", sheetID));
                DataTable dt = ado.loadDataTable(str, null, "vw_item");

                ContentPlaceHolder PlaceHolder2 = (ContentPlaceHolder)Page.Master.FindControl("ContentPlaceHolder2");
                PlaceHolder panel = (PlaceHolder)PlaceHolder2.FindControl("PH");

                //==================================================================================================
                //check(Machine/Location/Week/Operator)
                //==================================================================================================
                if (string.IsNullOrEmpty(txtMachine.Text)) { sb_NullOrEmpty.Append("【Machine No.】"); }
                if (string.IsNullOrEmpty(txtLocation.Text)) { sb_NullOrEmpty.Append("【Location】"); }
                if (string.IsNullOrEmpty(txtWeek.Text)) { sb_NullOrEmpty.Append("【Time/Week】"); }
                if (string.IsNullOrEmpty(txtOperator.Text)) { sb_NullOrEmpty.Append("【Operator】"); }


                List<string> arrStr = new List<string>();
                string transaction_Str1 = "", transaction_Str2 = "", transaction_Str3 = "", transaction_Str4 = "", transaction_Str5 = "", transaction_Str6 = "";
                foreach (DataRow dr in dt.Rows)
                {
                    //==================================================================================================
                    //檢查IsWeekly Or IsMonthly Item(Insert Mode)
                    //==================================================================================================
                    bool WeeklyOrMonthly = true;
                    if (string.IsNullOrEmpty(sheetID) && string.IsNullOrEmpty(Request.QueryString["preview"]))
                    {
                        if (dr["IsWeekly"].ToString().Equals("Y") && dr["IsMonthly"].ToString().Equals("N"))
                        {
                            int nowWeekID = Convert.ToInt32(func.getWeekCode(DateTime.Now));
                            int startWeekID = Convert.ToInt32(dr["StartWeek"].ToString());
                            int frequency = Convert.ToInt32(dr["frequency"].ToString());
                            if ((nowWeekID - startWeekID) % frequency != 0) { WeeklyOrMonthly = false; }
                        }

                        if (dr["IsMonthly"].ToString().Equals("Y") && dr["IsWeekly"].ToString().Equals("N"))
                        {
                            string getMonthID = func.getMonthCode(DateTime.Now);

                            //if no remainder, show items(check first week in this month?)
                            if (string.IsNullOrEmpty(getMonthID))
                            {
                                WeeklyOrMonthly = false;
                            }
                            else
                            {
                                int nowMonthID = Convert.ToInt32(getMonthID);
                                int startMonthID = Convert.ToInt32(dr["StartMonth"].ToString());
                                int Month_frequency = Convert.ToInt32(dr["Month_frequency"].ToString());

                                //if no remainder, show items(check first week in this month?)
                                if ((nowMonthID - startMonthID) % Month_frequency != 0) { WeeklyOrMonthly = false; }
                            }
                        }
                    }

                    if (WeeklyOrMonthly)
                    {
                        string item_value = string.Empty;
                        switch (dr["datatype_desc"].ToString())
                        {
                            case "TextBox":
                                TextBox txt = (TextBox)panel.FindControl(dr["itemID"].ToString());
                                item_value = txt.Text.ToString();
                                break;
                            case "RadioButton": //rdo_y_{0} rdo_n_{0} rdo_a_{0}
                                RadioButton rdo_y = (RadioButton)panel.FindControl(string.Format(@"rdo_y_{0}", dr["itemID"].ToString()));
                                RadioButton rdo_n = (RadioButton)panel.FindControl(string.Format(@"rdo_n_{0}", dr["itemID"].ToString()));
                                RadioButton rdo_a = (RadioButton)panel.FindControl(string.Format(@"rdo_a_{0}", dr["itemID"].ToString()));

                                if (rdo_y.Checked == false && rdo_n.Checked == false && rdo_a.Checked == false) { item_value = null; }
                                else
                                {
                                    if (rdo_y.Checked) item_value = "Y";
                                    else if (rdo_n.Checked) item_value = "N";
                                    else if (rdo_a.Checked) item_value = "A";
                                }
                                break;
                            //case "DropDownList":
                            //    DropDownList ddl = (DropDownList)panel.FindControl(dr["itemID"].ToString());
                            //    item_value = ddl.SelectedValue;
                            //    break;
                        }


                        //==================================================================================================
                        //check
                        //==================================================================================================
                        check(item_value, dr, CheckType.NullOrEmpty);
                        check(item_value, dr, CheckType.Validity);
                        check(item_value, dr, CheckType.Urgent);
                        check(item_value, dr, CheckType.Abnormal);


                        //==================================================================================================
                        //Insert / Update (sql_transaction)
                        //==================================================================================================                
                        string sqlStr = @"Select weekid From week Where :nowtime >start_date And :nowtime < end_date";
                        object[] param = new object[] { DateTime.Now, DateTime.Now };
                        DataTable wtb = ado.loadDataTable(sqlStr, param, "week");

                        //insert into sheet
                        transaction_Str1 = string.Format(@"insert into sheet(sheetID,sheet_CategoryID,describes,remarkid,weekid,Machine,Location,UsersID,InsertTime) 
                                                   values (sheet_sequence.nextval, '{0}' ,sheet_sequence.nextval,null,'{1}','{2}','{3}','{4}',SYSDATE)",
                                                                                    dr["sheet_categoryID"].ToString(), wtb.Rows[0]["weekid"].ToString(),
                                                                                    txtMachine.Text.Trim().ToUpper(), txtLocation.Text.Trim().ToUpper(), Session["UsersID"].ToString());
                        //insert into item
                        transaction_Str2 = string.Format(@"Insert into item(itemid,datatypeID,sheet_categoryID,sheetID,describes,
                                                                    method,spec,maxlimit,minlimit,value,iteminserttime,usersid,isUrgent,isNumeric,
                                                                    isWeekly,StartWeek,Frequency,
                                                                    isMonthly,StartMonth,Month_frequency)
                                                   Values (item_sequence.nextval,'{0}','{1}',sheet_sequence.currval,
                                                           '{2}','{3}','{4}','{5}','{6}',
                                                           '{7}',TO_DATE ('{8}','YYYY/MM/DD hh24:mi:ss'),'{9}',
                                                            '{10}','{11}',
                                                            '{12}','{13}','{14}',
                                                            '{15}','{16}','{17}')", dr["datatypeID"].ToString(), dr["sheet_categoryID"].ToString(),
                                                                                     dr["item_desc"].ToString(), dr["method"].ToString(), dr["spec"].ToString(),
                                                                                     dr["maxlimit"].ToString(), dr["minlimit"].ToString(), item_value,
                                                                                     Convert.ToDateTime(dr["iteminserttime"].ToString()).ToString("yyyy/MM/dd HH:mm:ss"), Session["UsersID"].ToString(),
                                                                                     dr["isUrgent"].ToString(), dr["IsNumeric"].ToString(),
                                                                                     dr["IsWeekly"].ToString(), dr["StartWeek"].ToString(), dr["Frequency"].ToString(),
                                                                                     dr["isMonthly"].ToString(), dr["StartMonth"].ToString(), dr["Month_frequency"].ToString());
                        //insert into Insert_delete_log
                        transaction_Str3 = string.Format(@"Insert into insert_delete_log (insert_delete_logID,sheetID,usersID,Log_action,log_Time)
                                                   Values (log_sequence.nextval,sheet_sequence.currval,'{0}','INSERT',sysdate)", Session["UsersID"].ToString());


                        //Update item
                        transaction_Str4 = string.Format(@"Update item
                                                   Set datatypeID='{0}',sheet_categoryID='{1}',sheetID='{2}',
                                                       describes='{3}',method='{4}',spec='{5}',maxlimit='{6}',minlimit='{7}',
                                                       value='{8}',iteminserttime=TO_DATE ('{9}','YYYY/MM/DD hh24:mi:ss'),usersid='{10}',
                                                       isUrgent='{11}',isNumeric='{12}',
                                                       IsWeekly='{13}',StartWeek='{14}',Frequency='{15}',
                                                       IsMonthly='{16}',StartMonth='{17}',Month_frequency='{18}'
                                                   Where itemid='{19}'", dr["datatypeID"].ToString(), dr["sheet_categoryID"].ToString(), sheetID,
                                                                               dr["item_desc"].ToString(), dr["method"].ToString(), dr["spec"].ToString(),
                                                                               dr["maxlimit"].ToString(), dr["minlimit"].ToString(), item_value,
                                                                               Convert.ToDateTime(dr["iteminserttime"].ToString()).ToString("yyyy/MM/dd HH:mm:ss"), Session["UsersID"].ToString(),
                                                                               dr["isUrgent"].ToString(), dr["IsNumeric"].ToString(),
                                                                               dr["IsWeekly"].ToString(), dr["StartWeek"].ToString(), dr["Frequency"].ToString(),
                                                                               dr["isMonthly"].ToString(), dr["StartMonth"].ToString(), dr["Month_frequency"].ToString(),
                                                                               dr["itemID"].ToString());

                        //insert into Update_log                
                        transaction_Str5 = string.Format(@"Insert into Update_Log(Update_LogID,itemID,usersID,oldvalue,newValue,updateTime)
                                                   Values (Update_Log_sequence.nextval,'{0}','{1}','{2}','{3}',sysdate)", dr["itemID"].ToString(),
                                                                                                                                 Session["UsersID"].ToString(),
                                                                                                                                 dr["value"].ToString(),
                                                                                                                                 item_value);

                        //Update sheet
                        transaction_Str6 = string.Format(@"Update sheet
                                                   Set UpdateTime=SYSDATE
                                                   Where sheetID='{0}'", sheetID);


                        if (string.IsNullOrEmpty(sheetID))//insert mode
                        {
                            arrStr.Add(transaction_Str2);
                        }
                        else//update mode
                        {
                            arrStr.Add(transaction_Str4);
                            if (!item_value.Equals(dr["value"].ToString())) { arrStr.Add(transaction_Str5); arrStr.Add(transaction_Str6); }
                        }
                    }
                }//end foreach

                if (string.IsNullOrEmpty(sheetID))//insert mode
                {
                    arrStr.Insert(0, transaction_Str1);
                    arrStr.Add(transaction_Str3);
                }

                if (!string.IsNullOrEmpty(sb_NullOrEmpty.ToString())) { sb_NullOrEmpty.Insert(0, "尚有欄位未填或未選擇\\n\\n"); }
                if (!string.IsNullOrEmpty(sb_Validity.ToString())) { sb_Validity.Insert(0, "\\n\\n資料輸入錯誤，請填入正確格式\\n\\n"); }
                if (!string.IsNullOrEmpty(sb_Urgent.ToString())) { sb_Urgent.Insert(0, "\\n\\n以下狀況請立即處理，並再做一次維修/調整後檢查\\n\\n"); }
                if (!string.IsNullOrEmpty(sb_Abnormal.ToString())) { sb_Abnormal.Insert(0, "\\n\\n欄位值不正常，已通知 Cell leader / engineers 前往處理\\n\\n"); }

                if (string.IsNullOrEmpty(sb_NullOrEmpty.ToString()) && string.IsNullOrEmpty(sb_Validity.ToString()) && string.IsNullOrEmpty(sb_Urgent.ToString()))
                {
                    string reStr = ado.SQL_transaction(arrStr, Conn);
                    if (reStr.ToUpper().Contains("SUCCESS"))
                    {
                        jsShow("alert('該筆資料新增成功。');");
                        btnReset_Click(null, null);
                    }
                    else
                        jsShow(string.Format(@"alert('新增失敗，該動作出現Exception。');"));


                    //==================================================================================================
                    //alert & mail
                    //================================================================================================== 
                    if (!string.IsNullOrEmpty(sb_Abnormal.ToString()))
                    {
                        string str_mail = @"SELECT sheetID,sheet_categoryid 
                                        FROM sheet 
                                        Order By sheetID Desc";
                        DataTable dt_mail = ado.loadDataTable(str_mail, null, "sheet");

                        StringBuilder sb_mail = new StringBuilder();

                        sb_mail.Append(@"<Html>
                                <head>
                                    <style type=text/css>
                                        .wrapper {border-collapse:collapse; font-family: verdana; font-size: 11px;}
                                        .DataTD_Green {text-align:left; background: #7E963D;color:white; padding: 1px 5px 1px 5px; border: 1px #C6D2DE solid; border-left: 1px #C6D2DEdotted; border-right: 1px #C6D2DE dotted; }
                                        .DataTD {text-align:left; background: #ffffff; padding: 1px 5px 1px 5px; border: 1px #C6D2DE solid; border-left: 1px #C6D2DE dotted;border-right: 1px #C6D2DE dotted; }
                                    </style>
                                </head>
                            <Body>");

                        sb_mail.Append(@"<p><b>欄位值不正常，已通知 Cell leader / engineers 前往處理</b></p>");
                        sb_mail.Append("<table class=wrapper align=left width=70% border=1>");
                        sb_mail.Append(string.Format(@"<tr><td class=DataTD_Green>檢視連結</td> <td class=DataTD>{0}View/eForm/sheet_Values.aspx?sheetID={1}&sheet_categoryID={2}</td></tr>",
                                                                                         ConfigurationManager.AppSettings["InternetURL"].ToString(),
                                                                                         dt_mail.Rows[0]["sheetID"].ToString(),
                                                                                         dt_mail.Rows[0]["sheet_categoryID"].ToString()));
                        sb_mail.Append(string.Format(@"<tr><td class=DataTD_Green>區域(Location)</td> <td class=DataTD>{0}</td></tr>", txtLocation.Text.Trim().ToUpper()));
                        sb_mail.Append(string.Format(@"<tr><td class=DataTD_Green>機台號碼(MachineNo)</td> <td class=DataTD>{0}</td></tr>", txtMachine.Text.Trim().ToUpper()));
                        sb_mail.Append(string.Format(@"<tr><td class=DataTD_Green>週別(Week Code)</td> <td class=DataTD>{0}</td></tr>", txtWeek.Text.Trim().ToUpper()));
                        sb_mail.Append(string.Format(@"<tr><td class=DataTD_Green>操作人員(Operator)</td> <td class=DataTD>{0}</td></tr>", txtOperator.Text.Trim().ToUpper()));
                        sb_mail.Append("</table>");
                        sb_mail.Append("<br /><p>");

                        sb_mail.Append(sb_Abnormal.ToString());
                        sb_mail.Append("</p><br />");

                        sb_mail.Append("<hr />");
                        sb_mail.Append("<p>******************ePM Mail Agent [ePM.mail.agent@tw-khh01.nxp.com]*******************</p>");
                        sb_mail.Append("</Body></Html>");


                        func.getMachineContactList(dt_mail.Rows[0]["sheet_categoryID"].ToString(), sb_mail);

                    }
                }
                else
                {
                    jsShow(string.Format(@"alert('{0}{1}{2}');", sb_NullOrEmpty.ToString(), sb_Validity.ToString(), sb_Urgent.ToString()));
                }
            }
            catch (Exception ex) 
            {
                string err = ex.ToString();    
            }
        }


        private DataTable getListItem(string scID)
        {
            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            DataTable dt = new DataTable();
            string str = string.Format(@"SELECT describes 
                                         FROM vw_ListItem
                                         WHERE sheet_categoryID='{0}'", scID);
            dt = ado.loadDataTable(str, null, "vw_ListItem");
            return dt;
        }
        private void jsShow(string str)
        {
            ScriptManager.RegisterStartupScript(this, this.GetType(), "js", str, true);
        }

        public enum CheckType : int
        { NullOrEmpty = 0, Validity = 1, Urgent = 2, Abnormal = 3 }
        private void check(string value, DataRow item_dr, CheckType type)
        {
            Common.ValidatorHelper val = new Common.ValidatorHelper();

            //describe of the item 
            string datatype_desc = item_dr["datatype_desc"].ToString();
            string item_desc = item_dr["item_desc"].ToString();
            string spec = item_dr["spec"].ToString();
            string method = item_dr["method"].ToString();
            string MaxLimit = item_dr["Maxlimit"].ToString();
            string MinLimit = item_dr["Minlimit"].ToString();
            string isUrgent = item_dr["isUrgent"].ToString();
            string isNumeric = item_dr["IsNumeric"].ToString();

            switch (type)
            {
                //==================================================================================================
                //Check NullOrEmpty
                //==================================================================================================
                case CheckType.NullOrEmpty:
                    switch (datatype_desc)
                    {
                        case "TextBox":
                            if (string.IsNullOrEmpty(value)) { sb_NullOrEmpty.Append(string.Format(@"【{0}】", item_desc)); }
                            break;
                        case "RadioButton":
                            if (string.IsNullOrEmpty(value)) { sb_NullOrEmpty.Append(string.Format(@"【{0}】", item_desc)); }
                            break;
                        //case "DropDownList":
                        //    if (value.Equals("00000")) { sb_NullOrEmpty.Append(string.Format(@"【{0}】", item_desc)); }
                        //    break;
                    }                    
                    break;
                //==================================================================================================
                //Check Validity
                //==================================================================================================
                case CheckType.Validity:
                    switch (datatype_desc)
                    {
                        case "TextBox":
                            if (isNumeric.Equals("N")) break;
                            if (string.IsNullOrEmpty(value)) break;
                            if (!val.IsFloat(value)) { sb_Validity.Append(string.Format(@"【{0}】", item_desc)); }
                            break;
                    }  
                    break;
                //==================================================================================================
                //Check Urgent
                //==================================================================================================
                case CheckType.Urgent:
                    if (isUrgent != "Y") break;
                    if (string.IsNullOrEmpty(value)) break;
                    switch (datatype_desc)
                    {
                        case "TextBox":
                            if (isNumeric.Equals("N")) break;
                            if (!val.IsFloat(value)) break;
                            if (!string.IsNullOrEmpty(MaxLimit) && (Convert.ToDouble(value) > Convert.ToDouble(MaxLimit)) ) 
                            {
                                sb_Urgent.Append(string.Format(@"【{0}，{1}】", item_desc, spec));
                                break;
                            }
                            if (!string.IsNullOrEmpty(MinLimit) && (Convert.ToDouble(value) < Convert.ToDouble(MinLimit)) )
                            {
                                sb_Urgent.Append(string.Format(@"【{0}，{1}】", item_desc, spec));
                                break;
                            }
                            break;
                        case "RadioButton":
                            if (value.Equals("A")) { sb_Urgent.Append(string.Format(@"【{0}，{1}】", item_desc, spec)); }
                            break;
                    }  
                    break;
                //==================================================================================================
                //Check Abnormal
                //==================================================================================================
                case CheckType.Abnormal:
                    if (isUrgent != "N") break;
                    if (string.IsNullOrEmpty(value)) break;
                    switch (datatype_desc)
                    {                            
                        case "TextBox":                            
                            if (isNumeric.Equals("N")) break;
                            if (!val.IsFloat(value)) break;
                            if (!string.IsNullOrEmpty(MaxLimit) && (Convert.ToDouble(value) > Convert.ToDouble(MaxLimit)))
                            {
                                sb_Abnormal.Append(string.Format(@"【{0}，{1}】", item_desc, spec));
                                break;
                            }
                            if (!string.IsNullOrEmpty(MinLimit) && (Convert.ToDouble(value) < Convert.ToDouble(MinLimit)))
                            {
                                sb_Abnormal.Append(string.Format(@"【{0}，{1}】", item_desc, spec));
                                break;
                            }
                            break;
                        case "RadioButton":
                            if (value.Equals("A")) { sb_Abnormal.Append(string.Format(@"【{0}，{1}】", item_desc, spec)); }
                            break;
                    }
                    break;
            }
        }

    }
}