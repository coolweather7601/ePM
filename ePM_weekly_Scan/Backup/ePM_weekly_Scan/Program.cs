using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Configuration;
using System.Data.OracleClient;

namespace EPM.Alan
{
    class Program
    {
        static void Main(string[] args)
        {
            string Conn = ePM_weekly_Scan.Properties.Settings.Default.EPM;

            //test schedule task project
            Common.AdoDbConn ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);

            //week_now
            string weekStr = @"select weekid from week where start_date <= :now_date And end_date >= :now_date";
            object[] para = new object[] { DateTime.Now, DateTime.Now }; ;
            DataTable dateTable = ado.loadDataTable(weekStr, para, "week");

            //log
            string logStr = @"Select Tester,Location 
                                  From ACS_Manage
                                  Where Tester not in (Select machine
                                                       From vw_insert_delete_log 
                                                       Where weekid = :now_weekid
                                                             And Log_Action ='INSERT') ";
            para = new object[] { dateTable.Rows[0]["weekid"].ToString() };
            DataTable logTable = ado.loadDataTable(logStr, para, "ACS_Manage");

            if (logTable.Rows.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(@"<table>
                               <tr><td colspan='2'>None Record Tester</td></tr>
                               <tr>
                                  <td>Tester</td>
                                  <td>Location</td>
                               </tr>");
                foreach (DataRow dr in logTable.Rows)
                {
                    sb.Append(string.Format(@"<td>{0}</td><td>{1}</td>", dr["Tester"].ToString(), dr["Location"].ToString()));
                }
                sb.Append("</table>");

                Common.MyFunc func = new Common.MyFunc();
                //func.mail("Alan.Kuo@nxp.com", sb.ToString());
            }
        }


    }
}
