using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.OracleClient;
using System.Text;
using System.Web.SessionState;
using System.Collections.Generic;

namespace EPM_Web.Alan.Common
{
    /// <summary>
    /// MyFunc 的摘要描述
    /// </summary>
    public class MyFunc
    {
        public static string ConnOpe2 = System.Configuration.ConfigurationManager.ConnectionStrings["ACS_ope2"].ToString();//ACS_ope2
        public static string ConnOpe1 = System.Configuration.ConfigurationManager.ConnectionStrings["ACS_ope1"].ToString();//ACS_ope1
        public static string HrConn = System.Configuration.ConfigurationManager.ConnectionStrings["HR"].ToString();//Hr        
        public static string MoldConn = System.Configuration.ConfigurationManager.ConnectionStrings["moldtool"].ToString();//MoldTool
        public static string EPMConn = System.Configuration.ConfigurationManager.ConnectionStrings["EPM"].ToString();//EPM

        public MyFunc()
        {
            //
            // TODO: 在此加入建構函式的程式碼
            //
        }

        public void mail(string contact, string mail_data)
        {
            string title = "ePM Check List Fail";
            using (OracleConnection connection = new OracleConnection(ConfigurationManager.ConnectionStrings["MAIL"].ConnectionString))
            {
                connection.Open();
                OracleCommand command = connection.CreateCommand();
                OracleTransaction transaction;
                transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);
                command.Transaction = transaction;
                try
                {
                    command.CommandText = "select seq_mail_id.nextval from dual";
                    int mail_id = Convert.ToInt32(command.ExecuteScalar());
                    command.CommandText = "insert into mail_pool (id,from_name,disable,datetime_in,send_period,mail_to,mail_cc,mail_subject,datetime_exp,exclusive_flag,check_sum,html_body) " +
                    "values (" + mail_id + ",'ePM mail agent',0,sysdate,0,'" + contact + "',null,'" + title + "',sysdate+1,0,null,1)";
                    command.ExecuteNonQuery();

                    command.CommandText = "insert into mail_body (id,sn,mail_cont) values (" + mail_id + ",1,'" + mail_data + "')";
                    command.ExecuteNonQuery();
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    //msg.Text = ex.ToString();
                }
            }
        }
        public object getAssignObj(string typeName, ControlCollection ctrs)
        {
            foreach (object r in ctrs)
            {
                string ss = r.GetType().Name.ToString();
                if (r.GetType().Name.ToString().Equals(typeName))
                    return r;
            }
            object null_obj = new object();
            return null_obj;
        }
        public DataTable readCsvTxt(string strpath)
        {
            int intColCount = 0;
            bool blnFlag = true;
            DataTable mydt = new DataTable("myTableName");

            DataColumn mydc;
            DataRow mydr;

            string strline;
            string[] aryline;

            System.IO.StreamReader mysr = new System.IO.StreamReader(strpath, Encoding.Default);

            while ((strline = mysr.ReadLine()) != null)
            {
                aryline = strline.Split(',');

                if (blnFlag)
                {
                    blnFlag = false;
                    intColCount = aryline.Length;
                    for (int i = 0; i < aryline.Length; i++)
                    {
                        mydc = new DataColumn(aryline[i]);
                        mydt.Columns.Add(mydc);
                    }
                }

                mydr = mydt.NewRow();
                for (int i = 0; i < intColCount; i++)
                {
                    mydr[i] = aryline[i];
                }
                mydt.Rows.Add(mydr);
            }

            return mydt;
        }
        public DataTable getAcsMachineData(MasterPage Master) //initial
        {
            string _conn = string.Empty, machineID = string.Empty;
            if (HttpContext.Current.Session["tester"] != null) { machineID = HttpContext.Current.Session["tester"].ToString(); }

            Common.AdoDbConn ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, EPMConn);
            DataTable dt = new DataTable();
            try
            {
                ContentPlaceHolder form = (ContentPlaceHolder)Master.FindControl("ContentPlaceHolder2");

                ((TextBox)form.FindControl("txtWeek")).Text = getWeekCode(DateTime.Now);

                string sql = string.Format(@"SELECT location
                                             FROM acs_manage
                                             WHERE tester = '{0}' AND location is not null", machineID);
                dt = ado.loadDataTable(sql, null, "acs_manage");
                if (dt.Rows.Count > 0)
                {
                    ((TextBox)form.FindControl("txtLocation")).Text = dt.Rows[0]["location"].ToString();
                    ((TextBox)form.FindControl("txtMachine")).Text = machineID;
                    ((TextBox)form.FindControl("txtOperator")).Text = HttpContext.Current.Session["account"].ToString();
                }
            }
            catch (Exception ce) { }
            return dt;
        }

        //==================================================================================================
        //Get data from webService
        //==================================================================================================
        public string getCgFromIntrack(string batchno)
        {
            LotData.SFCData sfc = new LotData.SFCData();
            LotData.CompleteLotDataProxy lot = sfc.getCompleteLotData(batchno);
            string CG = lot.AssemblyCG.ToString();
            return CG;
        }



        //==================================================================================================
        //Common check
        //==================================================================================================
        public enum Role : int
        {
            Administrator = 0, Supervisor = 1, User = 2
        }
        public void checkLogin()
        {
            if (HttpContext.Current.Session["account"] == null)
            {
                if (System.Configuration.ConfigurationManager.AppSettings["isDemo"].ToString().Equals("Y"))
                {
                    HttpContext.Current.Session["UsersID"] = "1";
                    HttpContext.Current.Session["account"] = "admin";
                    HttpContext.Current.Session["RoleID"] = "0";
                    HttpContext.Current.Session["Name"] = "系統管理者";
                    HttpContext.Current.Session["Dept_Name"] = "NXP";
                }
                else
                {
                    HttpContext.Current.Response.Redirect("../others/Login.aspx");
                }
            }
        }
        public void checkRole(MasterPage Master)
        {
            ContentPlaceHolder PlaceHolder1 = (ContentPlaceHolder)Master.FindControl("ContentPlaceHolder1");
            foreach (object obj in PlaceHolder1.Controls)
            {
                switch (obj.GetType().Name.ToString())
                {
                    #region MasterPage_Button
                    case "Button":
                        //==================================================================================
                        //btnRole
                        //==================================================================================
                        if (((Button)obj).ID.Equals("btnRole") && Convert.ToInt32(HttpContext.Current.Session["RoleID"]) == (int)Role.Administrator)
                        { ((Button)obj).Visible = true; }
                        break;
                    #endregion
                }
            }


            ContentPlaceHolder PlaceHolder2 = (ContentPlaceHolder)Master.FindControl("ContentPlaceHolder2");
            UpdatePanel panel = (UpdatePanel)PlaceHolder2.FindControl("up");
            Object ctrs = getAssignObj("Control", ((UpdatePanel)panel).Controls);

            if (ctrs.GetType().Name.ToString().Equals("Control"))
            {
                foreach (object ctr in ((Control)ctrs).Controls)
                {
                    switch (ctr.GetType().Name.ToString())
                    {
                        #region GridView
                        case "GridView":
                            string gvID = ((GridView)ctr).ID;

                            //==================================================================================
                            //ACS_Manage
                            //==================================================================================
                            if (gvID.Equals("GridViewAM") &&
                                (Convert.ToInt32(HttpContext.Current.Session["RoleID"]) == (int)Role.Administrator) || Convert.ToInt32(HttpContext.Current.Session["RoleID"]) == (int)Role.Supervisor)
                            {
                                ((GridView)ctr).Columns[4].Visible = true;
                                ((GridView)ctr).Columns[5].Visible = true;
                                ((GridView)ctr).Columns[6].Visible = true;
                            }


                            //==================================================================================
                            //List
                            //==================================================================================
                            if (gvID.Equals("GridViewList") &&
                                (Convert.ToInt32(HttpContext.Current.Session["RoleID"]) == (int)Role.Administrator || Convert.ToInt32(HttpContext.Current.Session["RoleID"]) == (int)Role.Supervisor))
                            {
                                //((GridView)ctr).Columns[8].Visible = true;
                                //((GridView)ctr).Columns[9].Visible = true;
                                //((GridView)ctr).Columns[10].Visible = true;
                                ((GridView)ctr).Columns[11].Visible = true;
                            }

                            //==================================================================================
                            //Sheet_category
                            //==================================================================================
                            if (gvID.Equals("GridViewSC") &&
                                (Convert.ToInt32(HttpContext.Current.Session["RoleID"]) == (int)Role.Administrator || Convert.ToInt32(HttpContext.Current.Session["RoleID"]) == (int)Role.Supervisor))
                            {
                                ((GridView)ctr).Columns[4].Visible = true;
                                ((GridView)ctr).Columns[5].Visible = true;
                                ((GridView)ctr).Columns[6].Visible = true;
                                ((GridView)ctr).Columns[7].Visible = true;
                                ((GridView)ctr).Columns[8].Visible = true;
                                ((GridView)ctr).Columns[9].Visible = true;
                            }
                            break;
                        #endregion

                        #region Buttton
                        case "Button":
                            //==================================================================================
                            //btnNew
                            //==================================================================================
                            if (((Button)ctr).ID.Equals("btnNew") &&
                                (Convert.ToInt32(HttpContext.Current.Session["RoleID"]) == (int)Role.Administrator || Convert.ToInt32(HttpContext.Current.Session["RoleID"]) == (int)Role.Supervisor))
                            { 
                                ((Button)ctr).Visible = true;
                                break;
                            }

                            //==================================================================================
                            //btnMail
                            //==================================================================================
                            if (((Button)ctr).ID.Equals("btnMail") && 
                                Convert.ToInt32(HttpContext.Current.Session["RoleID"]) == (int)Role.User)
                            { 
                                ((Button)ctr).Visible = false;
                                break;
                            }

                            //==================================================================================
                            //Demo
                            //==================================================================================
                            if (((Button)ctr).ID.Equals("btnSwitch") && 
                                Convert.ToInt32(HttpContext.Current.Session["RoleID"]) == (int)Role.Administrator &&
                                System.Configuration.ConfigurationManager.AppSettings["isDemo"].ToString().Equals("Y"))
                            { 
                                ((Button)ctr).Visible = true; 
                                break;
                            }
                            break;
                        #endregion
                    }
                }
            }

        }



        //==================================================================================================
        //Get data from Hr
        //==================================================================================================
        public DataTable getEmpDeptDt()
        {
            Common.AdoDbConn adoHr = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, HrConn);
            string deptStr = string.Format(@"SELECT distinct DEPT_NAME from EMP_ACCESS_LIST
                                             WHERE  (DEPT_NAME = 'ABE-CHEMICAL' OR DEPT_NAME = 'ABE-MOLDING' 
                                                     OR DEPT_NAME = 'ABE-SINGULATION' OR DEPT_NAME = 'AFE-DB' OR DEPT_NAME = 'AFE-WB'
                                                     OR DEPT_NAME = 'PA-PACKAGE SAW' OR DEPT_NAME = 'PA-PRE-ASSEMBLY' OR DEPT_NAME = 'WTF'
                                                     OR DEPT_NAME ='ABE-EED')
                                             AND email IS not NULL
                                             order by DEPT_NAME");
            DataTable dtDept = new DataTable();
            dtDept = adoHr.loadDataTable(deptStr, null, "EMP_ACCESS_LIST");
            return dtDept;
        }



        //==================================================================================================
        //Get data from DB
        //==================================================================================================
        public DataTable getDeptAllData(string deptName)
        {
            Common.AdoDbConn adoEPM = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, EPMConn);
            string empStr = string.Format(@"SELECT * FROM Users
                                         WHERE dept='{0}'
                                         AND mail IS NOT NULL
                                         ORDER BY name", deptName);
            DataTable dtEmp = new DataTable();
            dtEmp = adoEPM.loadDataTable(empStr, null, "Users");
            return dtEmp;
        }
        public DataTable getPersonalData(string empNo)
        {
            Common.AdoDbConn adoEPM = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, EPMConn);
            string str = string.Format(@"SELECT UsersID,EMP_NO,Name,Mail,Dept
                                         FROM Users
                                         WHERE EMP_NO='{0}'", empNo);
            DataTable dt = new DataTable();
            dt = adoEPM.loadDataTable(str, null, "Users");

            return dt;
        }
        public void getMachineContactList(string sheet_CategoryID, StringBuilder sb)
        {
            Common.AdoDbConn ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, EPMConn);
            DataTable dt = new DataTable();
            string sql = "SELECT mail FROM vw_mailContact WHERE sheet_CategoryID=:sheet_CategoryID";
            object[] para = new object[] { sheet_CategoryID };
            dt = ado.loadDataTable(sql, para, "vw_mailContact");

            string lst = "";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (i > 0) { lst += ";"; }
                lst += string.Format(@"{0}@nxp.com", dt.Rows[i]["mail"].ToString());
            }
            string[] aryMail = lst.Split(';');
            foreach (string email in aryMail)
            {
                mail(email, sb.ToString().Replace("\\n\\n", "<br/>"));
            }
        }



        //==================================================================================================
        //Get data from ACS
        //==================================================================================================    
        public string getTimeShift()
        {
            Common.AdoDbConn ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, ConnOpe2);
            DataTable dt = new DataTable();

            string str = "SELECT Shift_Start_Time FROM Shift_Table";
            dt = ado.loadDataTable(str, null, "Shift_Table");

            string[] day = dt.Rows[0]["Shift_Start_Time"].ToString().Split(':');
            string[] middle = dt.Rows[1]["Shift_Start_Time"].ToString().Split(':');
            string[] night = dt.Rows[2]["Shift_Start_Time"].ToString().Split(':');

            TimeSpan dayStart = new TimeSpan(Convert.ToInt32(day[0]), Convert.ToInt32(day[1]), 00);
            TimeSpan middleStart = new TimeSpan(Convert.ToInt32(middle[0]), Convert.ToInt32(middle[1]), 00); ;
            TimeSpan nightStart = new TimeSpan(Convert.ToInt32(night[0]), Convert.ToInt32(night[1]), 00);

            string timeShift = "";
            TimeSpan now = DateTime.Now.TimeOfDay;
            if (TimeSpan.Compare(now, dayStart) > 0 && TimeSpan.Compare(now, middleStart) < 0) { timeShift = "早班"; }
            if (TimeSpan.Compare(now, middleStart) > 0 && TimeSpan.Compare(now, nightStart) < 0) { timeShift = "中班"; }
            if (TimeSpan.Compare(now, nightStart) > 0 || TimeSpan.Compare(now, dayStart) < 0) { timeShift = "晚班"; }

            return timeShift;
        }
        public string getWeekCode(DateTime time)
        {
            Common.AdoDbConn ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, EPMConn);
            DataTable dt = new DataTable();

            string sqlStr = @"Select weekid From week Where :nowtime >= start_date And :nowtime < end_date";
            object[] param = new object[] { time, time };
            DataTable wtb = ado.loadDataTable(sqlStr, param, "week");

            return wtb.Rows[0]["weekid"].ToString().Substring(2,2);
        }
        public string getMonthCode(DateTime time)
        {
            Common.AdoDbConn ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, EPMConn);

            string str = string.Format(@"select * from (Select * from 
                                        (   
                                            Select * 
                                            From week 
                                            Where weekid > 1400
                                                  And  weekid < 1500
                                                  And month_id =(Select month_id From week Where :time >= start_date And :time < end_date)               
                                            Order by weekid
                                        )Where rownum <2)
                                        Where weekid = (Select weekid From week Where :time >= start_date And :time < end_date)
                                        ", Convert.ToInt32(time.Year.ToString().Substring(2, 2)) * 100
                                         , Convert.ToInt32(time.AddYears(1).Year.ToString().Substring(2, 2)) * 100);
            object[] param = new object[] { time, time, time, time };
            DataTable dt = ado.loadDataTable(str, param, "week");

            return (dt.Rows.Count > 0) ? dt.Rows[0]["month_id"].ToString() : null;
        }
    }
}