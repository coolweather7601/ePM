using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Configuration;
using System.Data.OracleClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace EPM.Alan
{
    class Program
    {
        static void Main(string[] args)
        {
            string Conn = ePM_weekly_Scan.Properties.Settings.Default.EPM;
            Common.AdoDbConn ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);

            //now
            string weekStr = @"select weekid from week where start_date <= :now_date And end_date >= :now_date";
            object[] para = new object[] { DateTime.Now, DateTime.Now }; ;
            System.Data.DataTable dateTable = ado.loadDataTable(weekStr, para, "week");

            #region Is Week Base OR Is Month Base(把不是的挑出來)
            string checkStr = string.Format(@"Select itemid,sheetid,isweekly,startweek,frequency,ismonthly,startmonth,month_frequency
                                        From item 
                                        Where (isdelete is null or isdelete ='N') and
                                              ((isweekly='Y' and startweek is not null) OR (ismonthly='Y' and startmonth is not null))
                                        Order by sheetid");
            System.Data.DataTable checkDt = ado.loadDataTable(checkStr, null, "item");
            string queryStr = "";
            foreach (System.Data.DataRow dr in checkDt.Rows)
            {
                if (dr["IsWeekly"].ToString().Equals("Y") && dr["IsMonthly"].ToString().Equals("N"))//weekly item
                {
                    int nowWeekID = Convert.ToInt32(dateTable.Rows[0]["weekid"].ToString().Substring(2, 2));
                    int startWeekID = Convert.ToInt32(dr["StartWeek"].ToString());
                    int frequency = Convert.ToInt32(dr["frequency"].ToString());

                    if ((nowWeekID - startWeekID) % frequency != 0) //if no remainder, show items
                    {
                        if (!queryStr.Contains(dr["sheetID"].ToString()))
                            queryStr += string.IsNullOrEmpty(queryStr) ? dr["sheetid"].ToString() : "," + dr["sheetid"].ToString();
                    }
                }
                else if (dr["IsMonthly"].ToString().Equals("Y") && dr["IsWeekly"].ToString().Equals("N"))//monthly item
                {
                    string monthstr = string.Format(@"select * from (Select * from 
                                        (   
                                            Select * 
                                            From week 
                                            Where weekid > {0}
                                                  And  weekid < {1}
                                                  And month_id =(Select month_id From week Where :time >= start_date And :time < end_date)               
                                            Order by weekid
                                        )Where rownum <2)
                                        Where weekid = (Select weekid From week Where :time >= start_date And :time < end_date)
                                        ", Convert.ToInt32(DateTime.Now.Year.ToString().Substring(2, 2)) * 100
                                         , Convert.ToInt32(DateTime.Now.AddYears(1).Year.ToString().Substring(2, 2)) * 100);
                    object[] monthparam = new object[] { DateTime.Now, DateTime.Now, DateTime.Now, DateTime.Now };
                    System.Data.DataTable monthdt = ado.loadDataTable(monthstr, monthparam, "week");
                    string getMonthID = monthdt.Rows.Count > 0 ? monthdt.Rows[0]["month_id"].ToString() : string.Empty;

                    //if no remainder, show items(check first week in this month?)
                    if (string.IsNullOrEmpty(getMonthID))
                    {
                        if (!queryStr.Contains(dr["sheetID"].ToString()))
                            queryStr += string.IsNullOrEmpty(queryStr) ? dr["sheetid"].ToString() : "," + dr["sheetid"].ToString();
                    }
                    else
                    {
                        int nowMonthID = Convert.ToInt32(getMonthID);
                        int startMonthID = Convert.ToInt32(dr["StartMonth"].ToString());
                        int Month_frequency = Convert.ToInt32(dr["Month_frequency"].ToString());

                        //if no remainder, show items(check first week in this month?)
                        if ((nowMonthID - startMonthID) % Month_frequency != 0)
                        {
                            if (!queryStr.Contains(dr["sheetID"].ToString()))
                                queryStr += string.IsNullOrEmpty(queryStr) ? dr["sheetid"].ToString() : "," + dr["sheetid"].ToString();
                        }
                    }
                }
            }

            #endregion

            //log
            string logStr = string.Format(@"Select Tester,Location,isSealing
                              From ACS_Manage
                              Where Tester not in (Select machine
                                                   From vw_insert_delete_log 
                                                   Where weekid = :now_weekid
                                                         And Log_Action ='INSERT'
                                                         {0}) ", string.IsNullOrEmpty(queryStr) ? string.Empty : string.Format(@" And sheetID not in ({0})", queryStr));
            para = new object[] { dateTable.Rows[0]["weekid"].ToString() };
            System.Data.DataTable logTable = new System.Data.DataTable();
            logTable = ado.loadDataTable(logStr, para, "ACS_Manage");

            //=================================================================================
            //Excel 2003-2007
            //=================================================================================
            #region initial
            //引用Excel Application類別
            _Application myExcel = null;

            //檢查PC有無Excel在執行
            //bool flag = false;
            //foreach (System.Diagnostics.Process item in Process.GetProcesses())
            //{
            //    if (item.ProcessName == "EXCEL")
            //    {
            //        flag = true;
            //        break;
            //    }
            //}
            //if (!flag)
            //{
            //    myExcel = new Microsoft.Office.Interop.Excel.Application();
            //}
            //else
            //{
            //    object obj = Marshal.GetActiveObject("Excel.Application");//引用已在執行的Excel
            //    myExcel = obj as Microsoft.Office.Interop.Excel.Application;
            //}

            //myExcel.Visible = false;//設false效能會比較好

            //引用Excel Application類別
            //_Application myExcel = null;
            //引用活頁簿類別 
            _Workbook myBook = null;
            //引用工作表類別
            _Worksheet mySheet = null;
            //引用Range類別 
            Range myRange = null;
            //開啟一個新的應用程式
            myExcel = new Microsoft.Office.Interop.Excel.Application();
            #endregion

            //加入新的活頁簿 
            myExcel.Workbooks.Add(true);
            //停用警告訊息
            myExcel.DisplayAlerts = false;
            //讓Excel文件可見 
            myExcel.Visible = true;
            //引用第一個活頁簿
            myBook = myExcel.Workbooks[1];
            //設定活頁簿焦點
            myBook.Activate();
            //引用第一個工作表
            mySheet = (_Worksheet)myBook.Worksheets[1];
            //命名工作表的名稱為 "Cells"
            mySheet.Name = "Cells";
            //設工作表焦點
            mySheet.Activate();

            #region sheet 1
            //Title
            myExcel.Cells[1, 1] = "Tester";
            myExcel.Cells[1, 2] = "Location";
            myExcel.Cells[1, 3] = "isSealing";
            myExcel.Cells[1, 4] = "Process";
            myExcel.Cells[1, 5] = "Model";

            myRange = (Range)mySheet.get_Range(myExcel.Cells[1, 1], myExcel.Cells[1, 1]);
            myRange.Select();
            //用陣列一次寫入資料 
            myRange.Columns.Cells.Interior.Color = "65535";
            myRange.Columns.Cells.Borders.ColorIndex = "0";
            //myRange.Columns.Cells.Borders.TintAndShade = "0";
            myRange.Columns.Cells.VerticalAlignment = Constants.xlCenter;

            myRange = (Range)mySheet.get_Range(myExcel.Cells[1, 2], myExcel.Cells[1, 2]);
            myRange.Select();
            myRange.Columns.Cells.Interior.Color = "5296274";
            myRange.Columns.Cells.Borders.ColorIndex = "0";
            myRange.Columns.Cells.VerticalAlignment = Constants.xlCenter;

            myRange = (Range)mySheet.get_Range(myExcel.Cells[1, 3], myExcel.Cells[1, 3]);
            myRange.Select();
            myRange.Columns.Cells.Interior.Color = "49407";
            myRange.Columns.Cells.Borders.ColorIndex = "0";
            myRange.Columns.Cells.VerticalAlignment = Constants.xlCenter;

            myRange = (Range)mySheet.get_Range(myExcel.Cells[1, 4], myExcel.Cells[1, 4]);
            myRange.Select();
            myRange.Columns.Cells.Interior.Color = "12611584";
            myRange.Columns.Cells.Borders.ColorIndex = "0";
            myRange.Columns.Cells.VerticalAlignment = Constants.xlCenter;

            myRange = (Range)mySheet.get_Range(myExcel.Cells[1, 5], myExcel.Cells[1, 5]);
            myRange.Select();
            myRange.Columns.Cells.Interior.Color = "15773696";
            myRange.Columns.Cells.Borders.ColorIndex = "0";
            myRange.Columns.Cells.VerticalAlignment = Constants.xlCenter;

            int i = 0;
            foreach (System.Data.DataRow dr in logTable.Rows)
            {
                string acsConn = ePM_weekly_Scan.Properties.Settings.Default.ACS;
                Common.AdoDbConn acsAdo = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, acsConn);

                string acsStr = string.Format(@"Select * 
                                                From VW_Machine_List_ALL
                                                Where Machine='{0}'", dr["Tester"].ToString().ToUpper());
                System.Data.DataTable acsDt = acsAdo.loadDataTable(acsStr, null, "VW_Machine_List_ALL");

                myExcel.Cells[2 + i, 1] = "'" + dr["Tester"].ToString();
                myExcel.Cells[2 + i, 2] = "'" + dr["Location"].ToString();
                myExcel.Cells[2 + i, 3] = "'" + dr["isSealing"].ToString();

                if (acsDt.Rows.Count > 0)
                {
                    myExcel.Cells[2 + i, 4] = "'" + acsDt.Rows[0]["PROCESS"].ToString();
                    myExcel.Cells[2 + i, 5] = "'" + acsDt.Rows[0]["MODEL"].ToString();
                }

                if (dr["isSealing"].ToString().Equals("Y"))
                {
                    myRange = (Range)mySheet.get_Range(myExcel.Cells[2 + i, 1], myExcel.Cells[2 + i, 3]);
                    myRange.Select();
                    myRange.Columns.Cells.Interior.Color = "255";
                    myRange.Columns.Cells.Borders.ColorIndex = "0";
                }
                i++;
            }

            //myRange = (Range)mySheet.get_Range(myExcel.Cells[2, 1], myExcel.Cells[2 + (i-1), 2]);
            //myRange.Select();
            //myRange.Columns.Cells.Interior.Color = "5296274";
            //myRange.Columns.Cells.Borders.ColorIndex = "0";
            //myRange.Columns.Cells.VerticalAlignment = Constants.xlCenter;
            #endregion
            #region sheet 2
            ////加入新的工作表在第1張工作表之後
            //myBook.Sheets.Add(Type.Missing, myBook.Worksheets[1], 1, Type.Missing);
            ////引用第2個工作表
            //mySheet = (_Worksheet)myBook.Worksheets[2];
            ////命名工作表的名稱為 "Array"
            //mySheet.Name = "Array";//加入新的工作表在第1張工作表之後 
            //myBook.Sheets.Add(Type.Missing, myBook.Worksheets[1], 1, Type.Missing);
            ////引用第2個工作表 
            //mySheet = (_Worksheet)myBook.Worksheets[2];
            ////命名工作表的名稱為 "Array" 
            //mySheet.Name = "Array2";
            ////寫入報表名稱 
            //myExcel.Cells[1, 4] = "普通報表";
            ////設定範圍 
            //myRange = (Range)mySheet.get_Range(myExcel.Cells[2, 2], myExcel.Cells[4, 8]);
            //myRange.Select();
            ////用陣列一次寫入資料 
            //myRange.Value2 = "'test'";
            #endregion

            //設定儲存路徑 
            string FileName = string.Format(@"{0}_ePM_NoneRecordTester.xls", DateTime.Now.ToString("yyyy-MM-dd"));
            string dir = ePM_weekly_Scan.Properties.Settings.Default.websiteDir;
            string PathFile = dir + FileName;
            //string PathFile = Directory.GetCurrentDirectory() + string.Format(@"\{0}_ePM_NoneRecordTester.xls", DateTime.Now.ToString("yyyy-MM-dd"));
            //另存活頁簿 
            myBook.SaveAs(PathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
                                        , XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //關閉活頁簿 
            myBook.Close(false, Type.Missing, Type.Missing);
            //關閉Excel 
            myExcel.Quit();
            //釋放Excel資源 
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myExcel);
            myBook = null;
            mySheet = null;
            myRange = null;
            myExcel = null;
            GC.Collect();

            //=================================================================================
            //Mail
            //=================================================================================            
            if (logTable.Rows.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                string webUrl = ePM_weekly_Scan.Properties.Settings.Default.webSiteUrl;
                string downloadUrl = webUrl + FileName;

                sb.Insert(0, @"<p><b>Week</b></p>");
                sb.Insert(0, @"<Html>
                                <head>
                                    <style type=text/css>
                                        .wrapper {border-collapse:collapse; font-family: verdana; font-size: 11px;}
                                        .DataTD_Green {text-align:left; background: #7E963D;color:white; padding: 1px 5px 1px 5px; border: 1px #C6D2DE solid; border-left: 1px #C6D2DEdotted; border-right: 1px #C6D2DE dotted; }
                                        .DataTD {text-align:left; background: #ffffff; padding: 1px 5px 1px 5px; border: 1px #C6D2DE solid; border-left: 1px #C6D2DE dotted;border-right: 1px #C6D2DE dotted; }
                                    </style>
                                </head>
                            <Body>");

                sb.Append("<table class=wrapper align=left width=70% border=1>");
                sb.Append(string.Format(@"<tr>
                                              <td class=DataTD_Green>Reporting</td>
                                              <td class=DataTD>{0}</td>
                                         </tr>", downloadUrl));
                sb.Append("</table>");
                sb.Append("<br/><p>******************ePM Check List Mail Agent [ePM.mail.agent@tw-khh01.nxp.com]*******************</p>");
                sb.Append("</Body></Html>");


                Common.MyFunc func = new Common.MyFunc();

                //ABE-Molding:
                //PA-PRE-ASSEMBLY and PA-PACKAGE SAW:
                //AFE-DB:
                //AFE-WB:
                //ABE-CHEMICAL + ABE-SINGULATION:
                string[] contact_List = new string[] { "Abem.ts@nxp.com","kc.kung-chiao.chen@nxp.com","Sun.lin@nxp.com","c.c.hunag@nxp.com",
                                                       "K.H.Chen@nxp.com","T.S.MLPA@nxp.com","WTFSG.RM@nxp.com","Lion.ho@nxp.com",
                                                       "Gary.huang@nxp.com","Y.h.wu@nxp.com","chien-chou.kuo@nxp.com","db.rm@nxp.com",
                                                       "C.J.hsueh@nxp.com","WB.rm@nxp.com","K.F.Fu@nxp.com",
                                                       "i.c.tsai@nxp.com","terry.tl.wu@nxp.com","c.h.chang@nxp.com","m.l.sun@nxp.com","h.c.wei@nxp.com",
                                                       "sj.chen@nxp.com","abec.ts@nxp.com","abes.ts@nxp.com","wayne.huang@nxp.com","rudy.chen@nxp.com",
                                                       "Alan.Kuo@nxp.com"
                                                     };

                foreach (string contact in contact_List)
                {
                    func.mail(contact, sb.ToString());
                }
            }
        }
    }
}
