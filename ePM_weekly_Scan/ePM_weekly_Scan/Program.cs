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
//NOPI
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

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


            #region Is Week Base OR Is Month Base(pick up that sheetid should record this week )
            string allStr = string.Format(@" Select sheet_categoryid,Tester 
                                             From sheet_category
                                             Where (isdelete is null or isdelete = 'N') ");
            System.Data.DataTable allDt = ado.loadDataTable(allStr, null, "sheet_Category");

            string queryStr = "";
            string abnormal_queryStr = "";
            foreach (System.Data.DataRow dr in allDt.Rows)
            {
                //============================================================================================================
                //take the last setting, check if is's abnormal showit in other sheet(check)
                //============================================================================================================
                string checkStr = string.Format(@"select * from(
                                                select count(*) as row_count,isweekly,startweek,frequency,ismonthly,startmonth,month_frequency
                                                from item
                                                where sheet_categoryid='{0}'
                                                and sheetid is null
                                                and (isweekly='N' and ismonthly='Y') and (isdelete is null or isdelete = 'N')
                                                group by isweekly,startweek,frequency,ismonthly,startmonth,month_frequency)
                                                union 
                                            (
                                                select count(*) as YN_count,isweekly as NN_isweekly,startweek as NN_startweek,frequency as NN_frequency,ismonthly as NN_ismonthly,startmonth as NN_startmonth,month_frequency as NN_month_frequency
                                                from item
                                                where sheet_categoryid='{0}'
                                                    and sheetid is null
                                                    and (isweekly='Y' and ismonthly='N') and (isdelete is null or isdelete = 'N')
                                                group by isweekly,startweek,frequency,ismonthly,startmonth,month_frequency)
                                                union    
                                            (
                                                select count(*) as NN_count,isweekly as NN_isweekly,startweek as NN_startweek,frequency as NN_frequency,ismonthly as NN_ismonthly,startmonth as NN_startmonth,month_frequency as NN_month_frequency
                                                from item
                                                where sheet_categoryid='{0}'
                                                    and sheetid is null
                                                    and (isweekly='N' and ismonthly='N') and (isdelete is null or isdelete = 'N')
                                                group by isweekly,startweek,frequency,ismonthly,startmonth,month_frequency
                                            )", dr["sheet_categoryid"].ToString());
                System.Data.DataTable checkDt = ado.loadDataTable(checkStr, null, "item");

                string _startWeek = "", _startMonth = "", _frequency = "", _month_frequency = "";
                foreach (System.Data.DataRow checkdr in checkDt.Rows)
                {
                    //============================================================================================================
                    //dayly item(NN)
                    //============================================================================================================
                    if (checkdr["isweekly"].ToString().Equals("N") && checkdr["ismonthly"].ToString().Equals("N"))
                    {
                        if (!string.IsNullOrEmpty(dr["tester"].ToString()))
                        {
                            string[] mArrs = dr["tester"].ToString().Split('|');
                            foreach (string m in mArrs)
                            {
                                if (!string.IsNullOrEmpty(m))
                                {
                                    string nn_Str = string.Format(@"select tester from acs_manage where acs_manageid='{0}'", m);
                                    System.Data.DataTable nnTable = ado.loadDataTable(nn_Str, null, "acs_manage");
                                    if (nnTable.Rows.Count > 0) queryStr = string.IsNullOrEmpty(queryStr) ? string.Format(@" (machine='{0}' )", nnTable.Rows[0]["tester"].ToString()) : string.Format(@"{0} Or machine='{1}')", queryStr.Replace(")", ""), nnTable.Rows[0]["tester"].ToString());
                                }
                            }
                        }
                        //queryStr = string.IsNullOrEmpty(queryStr) ? string.Format(@" (sheetid='{0}' )", dr["sheetid"].ToString()) : string.Format(@"{0} Or sheetid='{1}')", queryStr.Replace(")", ""), dr["sheetid"].ToString());
                    }
                    //============================================================================================================
                    //weekly item(YN) weekly item
                    //============================================================================================================
                    else if (checkdr["isweekly"].ToString().Equals("Y") && checkdr["ismonthly"].ToString().Equals("N"))
                    {
                        _startWeek = checkdr["startWeek"].ToString();
                        _frequency = checkdr["frequency"].ToString();

                        int nowWeekID = Convert.ToInt32(dateTable.Rows[0]["weekid"].ToString().Substring(2, 2));
                        int startWeekID = Convert.ToInt32(_startWeek);
                        int frequency = Convert.ToInt32(_frequency);

                        if ((nowWeekID - startWeekID) % frequency == 0) //if no remainder, show items
                        {
                            if (!string.IsNullOrEmpty(dr["tester"].ToString()))
                            {
                                string[] mArrs = dr["tester"].ToString().Split('|');
                                foreach (string m in mArrs)
                                {
                                    if (!string.IsNullOrEmpty(m))
                                    {
                                        string yn_Str = string.Format(@"select tester from acs_manage where acs_manageid='{0}'", m);
                                        System.Data.DataTable ynTable = ado.loadDataTable(yn_Str, null, "acs_manage");
                                        if (ynTable.Rows.Count > 0) queryStr = string.IsNullOrEmpty(queryStr) ? string.Format(@" (machine='{0}' )", ynTable.Rows[0]["tester"].ToString()) : string.Format(@"{0} Or machine='{1}')", queryStr.Replace(")", ""), ynTable.Rows[0]["tester"].ToString());
                                    }
                                }
                            }
                            //queryStr = string.IsNullOrEmpty(queryStr) ? string.Format(@" (sheetid='{0}' )", dr["sheetid"].ToString()) : string.Format(@"{0} Or sheetid='{1}')", queryStr.Replace(")", ""), dr["sheetid"].ToString());

                        }
                    }
                    //============================================================================================================
                    //monthly item(NY) monthly item
                    //============================================================================================================
                    else if (checkdr["isweekly"].ToString().Equals("N") && checkdr["ismonthly"].ToString().Equals("Y"))
                    {
                        _startMonth = checkdr["startMonth"].ToString();
                        _month_frequency = checkdr["month_frequency"].ToString();

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

                        //if no remainder, show items(check first week in this month?)
                        if (monthdt.Rows.Count > 0)
                        {
                            int nowMonthID = Convert.ToInt32(monthdt.Rows[0]["month_id"].ToString());
                            int startMonthID = Convert.ToInt32(_startMonth);
                            int Month_frequency = Convert.ToInt32(_month_frequency);

                            //if no remainder, show items(check first week in this month?)
                            if ((nowMonthID - startMonthID) % Month_frequency == 0)
                            {
                                if (!string.IsNullOrEmpty(dr["tester"].ToString()))
                                {
                                    string[] mArrs = dr["tester"].ToString().Split('|');
                                    foreach (string m in mArrs)
                                    {
                                        if (!string.IsNullOrEmpty(m))
                                        {
                                            string ny_Str = string.Format(@"select tester from acs_manage where acs_manageid='{0}'", m);
                                            System.Data.DataTable nyTable = ado.loadDataTable(ny_Str, null, "acs_manage");
                                            if (nyTable.Rows.Count > 0) queryStr = string.IsNullOrEmpty(queryStr) ? string.Format(@" (machine='{0}' )", nyTable.Rows[0]["tester"].ToString()) : string.Format(@"{0} Or machine='{1}')", queryStr.Replace(")", ""), nyTable.Rows[0]["tester"].ToString());
                                        }
                                    }
                                }
                                //queryStr = string.IsNullOrEmpty(queryStr) ? string.Format(@" (sheetid='{0}' )", dr["sheetid"].ToString()) : string.Format(@"{0} Or sheetid='{1}')", queryStr.Replace(")", ""), dr["sheetid"].ToString());
                            }
                        }
                    }
                    //============================================================================================================
                    //error setting
                    //============================================================================================================
                    else if (checkdr["isweekly"].ToString().Equals("Y") && checkdr["ismonthly"].ToString().Equals("Y"))
                    {
                        if (!string.IsNullOrEmpty(dr["sheet_categoryid"].ToString()))
                            abnormal_queryStr = string.IsNullOrEmpty(abnormal_queryStr) ? string.Format(@" (sheet_categoryid='{0}' )", dr["sheet_categoryid"].ToString()) : string.Format(@"{0} Or sheet_categoryid='{1}')", abnormal_queryStr.Replace(")", ""), dr["sheet_categoryid"].ToString());
                    }
                }
            }//end foreach
            #endregion

            //abnormal
            string ab_Str = string.Format(@"select sheet_categoryid,describes,docnumber
                                            from sheet_category 
                                            where (isdelete is null or isdelete='N') {0}
                                            ", string.IsNullOrEmpty(abnormal_queryStr) ?
                                                    string.Format(@" And tester='NA'") :
                                                    string.Format(@" And {0}", abnormal_queryStr));
            System.Data.DataTable ab_logTable = new System.Data.DataTable();
            ab_logTable = ado.loadDataTable(ab_Str, null, "sheet_category");

            //log_1(test)
            string logStr_1 = string.Format(@"Select machine
                                              From vw_insert_delete_log 
                                              Where weekid = :now_weekid
                                                 And Log_Action ='INSERT'
                                                 {0}
                                            ", string.IsNullOrEmpty(queryStr) ? string.Empty : string.Format(@" And {0}", queryStr), queryStr);

            para = new object[] { dateTable.Rows[0]["weekid"].ToString() };
            System.Data.DataTable logTable_1 = new System.Data.DataTable();
            logTable_1 = ado.loadDataTable(logStr_1, para, "vw_insert_delete_log");

            //log_2(test)
            string logStr_2 = string.Format(@"Select machine 
                                              From vw_insert_delete_log
                                              Where {0} 
                                              group by machine
                                            ", queryStr);
            System.Data.DataTable logTable_2 = new System.Data.DataTable();
            logTable_2 = ado.loadDataTable(logStr_2, para, "vw_insert_delete_log");

            //log
            string logStr = string.Format(@"Select Tester,Location,isSealing
                                            From ACS_Manage
                                            Where Tester not in (Select machine
                                                                 From vw_insert_delete_log 
                                                                 Where weekid = :now_weekid
                                                                 And Log_Action ='INSERT'
                                                                 {0}) 
                                                  And Tester in (Select machine 
                                                                 From vw_insert_delete_log
                                                                 Where {1} group by machine)
                                            ", string.IsNullOrEmpty(queryStr) ? string.Empty : string.Format(@" And {0}", queryStr), queryStr);

            para = new object[] { dateTable.Rows[0]["weekid"].ToString() };
            System.Data.DataTable logTable = new System.Data.DataTable();
            logTable = ado.loadDataTable(logStr, para, "ACS_Manage");

            //=================================================================================
            //Excel
            //=================================================================================

            string title = "None Record Machines";
            string title_2 = "Error Setting List";
            HSSFWorkbook hssfWorkBook_1 = new HSSFWorkbook();

            #region sheet1
            HSSFSheet sheet = (NPOI.HSSF.UserModel.HSSFSheet)hssfWorkBook_1.CreateSheet(title);
            IRow row; ICell cell;
            bool isOnly = false;
            int Row_Count = 0;
            int Cell_Count = 0;


            row = sheet.CreateRow(Row_Count);
            string[] xls_title = new string[] { "Tester", "Location", "isSealing", "Process", "Model" };
            short[] colorLst = new short[] { NPOI.HSSF.Util.HSSFColor.YELLOW.index2, NPOI.HSSF.Util.HSSFColor.TEAL.index2, NPOI.HSSF.Util.HSSFColor.SKY_BLUE.index, NPOI.HSSF.Util.HSSFColor.ROSE.index, NPOI.HSSF.Util.HSSFColor.ORANGE.index };
            foreach (string _title in xls_title)
            {
                cell = row.CreateCell(Cell_Count);
                cell.SetCellValue(_title.ToString());
                cell.CellStyle.FillBackgroundColor = colorLst[Cell_Count];
                Cell_Count++;
            }
            if (isOnly == false)
            {
                sheet.SetAutoFilter(CellRangeAddress.ValueOf(string.Format(@"A{0}:E{0}", Row_Count + 1)));//Fliter
                sheet.CreateFreezePane(Cell_Count, Row_Count + 1);//Freeze
                isOnly = true;
            }
            Row_Count += 1;

            //set value
            foreach (DataRow dr in logTable.Rows)
            {
                Cell_Count = 0;
                row = sheet.CreateRow(Row_Count);

                string acsConn = ePM_weekly_Scan.Properties.Settings.Default.ACS;
                Common.AdoDbConn acsAdo = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, acsConn);

                string acsStr = string.Format(@"Select * 
                                                From VW_Machine_List_ALL
                                                Where Machine='{0}'", dr["Tester"].ToString().ToUpper());
                System.Data.DataTable acsDt = acsAdo.loadDataTable(acsStr, null, "VW_Machine_List_ALL");

                //Tester
                cell = row.CreateCell(Cell_Count);
                cell.SetCellValue(dr["Tester"].ToString());
                Cell_Count++;

                //Location
                cell = row.CreateCell(Cell_Count);
                cell.SetCellValue(dr["Location"].ToString());
                Cell_Count++;

                //isSealing
                cell = row.CreateCell(Cell_Count);
                cell.SetCellValue(dr["isSealing"].ToString());
                Cell_Count++;

                if (acsDt.Rows.Count > 0)
                {
                    //Process
                    cell = row.CreateCell(Cell_Count);
                    cell.SetCellValue(acsDt.Rows[0]["Process"].ToString());
                    Cell_Count++;

                    //Model
                    cell = row.CreateCell(Cell_Count);
                    cell.SetCellValue(acsDt.Rows[0]["Model"].ToString());
                    Cell_Count++;
                }

                if (dr["isSealing"].ToString().Equals("Y"))// color red
                {
                    //myRange = (Range)mySheet.get_Range(myExcel.Cells[2 + i, 1], myExcel.Cells[2 + i, 3]);
                    //myRange.Select();
                    //myRange.Columns.Cells.Interior.Color = "255";
                    //myRange.Columns.Cells.Borders.ColorIndex = "0";
                }

                Row_Count++;
            }
            #endregion
            #region sheet2
            HSSFSheet sheet_2 = (NPOI.HSSF.UserModel.HSSFSheet)hssfWorkBook_1.CreateSheet(title_2);
            IRow row_2; ICell cell_2;
            bool isOnly_2 = false;
            int Row_Count_2 = 0;
            int Cell_Count_2 = 0;

            row_2 = sheet_2.CreateRow(Row_Count_2);
            string[] xls_title_2 = new string[] { "sheet_categoryID", "Category_Name", "docnumber" };
            short[] colorLst_2 = new short[] { NPOI.HSSF.Util.HSSFColor.YELLOW.index2, NPOI.HSSF.Util.HSSFColor.TEAL.index2, NPOI.HSSF.Util.HSSFColor.SKY_BLUE.index };
            foreach (string _title in xls_title_2)
            {
                cell_2 = row_2.CreateCell(Cell_Count_2);
                cell_2.SetCellValue(_title.ToString());
                cell_2.CellStyle.FillBackgroundColor = colorLst_2[Cell_Count_2];
                Cell_Count_2++;
            }
            if (isOnly == false)
            {
                sheet_2.SetAutoFilter(CellRangeAddress.ValueOf(string.Format(@"A{0}:C{0}", Row_Count_2 + 1)));//Fliter
                sheet_2.CreateFreezePane(Cell_Count_2, Row_Count_2 + 1);//Freeze
                isOnly_2 = true;
            }
            Row_Count_2 += 1;

            //set value
            foreach (DataRow dr in ab_logTable.Rows)
            {
                Cell_Count_2 = 0;
                row_2 = sheet_2.CreateRow(Row_Count_2);

                //sheet_categoryid
                cell_2 = row_2.CreateCell(Cell_Count_2);
                cell_2.SetCellValue(dr["sheet_categoryid"].ToString());
                Cell_Count_2++;

                //describes
                cell_2 = row_2.CreateCell(Cell_Count_2);
                cell_2.SetCellValue(dr["describes"].ToString());
                Cell_Count_2++;

                //docnumber
                cell_2 = row_2.CreateCell(Cell_Count_2);
                cell_2.SetCellValue(dr["docnumber"].ToString());
                Cell_Count_2++;

                Row_Count_2++;
            }
            #endregion

            //Save File (\download\..)
            string FileName = string.Format(@"{0}_ePM__NoneRecordTester.xls", DateTime.Now.ToString("yyyy-MM-dd"));
            string dir = ePM_weekly_Scan.Properties.Settings.Default.websiteDir;
            string PathFile = dir + FileName;
            FileStream fs = new FileStream(PathFile, FileMode.Create);
            hssfWorkBook_1.Write(fs);
            fs.Close();
            fs.Dispose();


            //=================================================================================
            //Mail
            //=================================================================================            
            if (logTable.Rows.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                string webUrl = ePM_weekly_Scan.Properties.Settings.Default.webSiteUrl;
                string downloadUrl = webUrl + FileName;

                sb.Insert(0, string.Format(@"<p><b>Week {0}</b></p>", dateTable.Rows[0]["weekid"].ToString().Substring(1, 3)));
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
                                                       "Alan.Kuo@nxp.com", "t.w.lin@nxp.com"
                                                     };
                string mailList = "";
                foreach (string contact in contact_List)
                {
                    if (!string.IsNullOrEmpty(mailList)) { mailList += ","; }
                    mailList += contact;
                }
                func.mail(mailList, sb.ToString());
            }
        }

        static void Main2(string[] args)
        {
            string Conn = ePM_weekly_Scan.Properties.Settings.Default.EPM;
            Common.AdoDbConn ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);

            //now
            string weekStr = @"select weekid from week where start_date <= :now_date And end_date >= :now_date";
            object[] para = new object[] { DateTime.Now, DateTime.Now }; ;
            System.Data.DataTable dateTable = ado.loadDataTable(weekStr, para, "week");


            #region Is Week Base OR Is Month Base(pick up that sheetid should record this week )
            string allStr = string.Format(@" Select sheet_categoryid,sheetid
                                             From item 
                                             Where (isdelete is null or isdelete ='N') 
                                                 and ((isweekly='Y' and startweek is not null) OR (ismonthly='Y' and startmonth is not null) or (ismonthly='N' and isweekly='N'))
                                                  and sheetid is not null
                                             group by sheet_categoryid,sheetid                                              
                                             Order by sheetid");
            System.Data.DataTable allDt = ado.loadDataTable(allStr, null, "item");

            string queryStr = "";
            string abnormal_queryStr = "";
            foreach (System.Data.DataRow dr in allDt.Rows)
            {
                //============================================================================================================
                //take the last setting, check if is's abnormal showit in other sheet(check)
                //============================================================================================================
                string checkStr = string.Format(@"select * from(
                                                select count(*) as row_count,isweekly,startweek,frequency,ismonthly,startmonth,month_frequency
                                                from item
                                                where sheet_categoryid='{0}'
                                                and sheetid is null
                                                and (isweekly='N' and ismonthly='Y')
                                                group by isweekly,startweek,frequency,ismonthly,startmonth,month_frequency)
                                                union 
                                            (
                                                select count(*) as YN_count,isweekly as NN_isweekly,startweek as NN_startweek,frequency as NN_frequency,ismonthly as NN_ismonthly,startmonth as NN_startmonth,month_frequency as NN_month_frequency
                                                from item
                                                where sheet_categoryid='{0}'
                                                    and sheetid is null
                                                    and (isweekly='Y' and ismonthly='N')
                                                group by isweekly,startweek,frequency,ismonthly,startmonth,month_frequency)
                                                union    
                                            (
                                                select count(*) as NN_count,isweekly as NN_isweekly,startweek as NN_startweek,frequency as NN_frequency,ismonthly as NN_ismonthly,startmonth as NN_startmonth,month_frequency as NN_month_frequency
                                                from item
                                                where sheet_categoryid='{0}'
                                                    and sheetid is null
                                                    and (isweekly='N' and ismonthly='N')
                                                group by isweekly,startweek,frequency,ismonthly,startmonth,month_frequency
                                            )", dr["sheet_categoryid"].ToString());
                System.Data.DataTable checkDt = ado.loadDataTable(checkStr, null, "item");

                bool NNcheck = false, YNcheck = false, NYcheck = false;
                string _startWeek = "", _startMonth = "", _frequency = "", _month_frequency = "";
                foreach (System.Data.DataRow checkdr in checkDt.Rows)
                {
                    if (checkdr["isweekly"].ToString().Equals("N") && checkdr["ismonthly"].ToString().Equals("N")) { NNcheck = true; }
                    else if (checkdr["isweekly"].ToString().Equals("Y") && checkdr["ismonthly"].ToString().Equals("N")) { YNcheck = true; _startWeek = checkdr["startWeek"].ToString(); _frequency = checkdr["frequency"].ToString(); }
                    //else if (checkdr["isweekly"].ToString().Equals("Y") && checkdr["ismonthly"].ToString().Equals("N") && (checkdr["startweek"].ToString() != "1" && checkdr["frequency"].ToString() != "1")) { YNcheck = true; _startWeek = checkdr["startWeek"].ToString(); _frequency = checkdr["frequency"].ToString(); }
                    else if (checkdr["isweekly"].ToString().Equals("N") && checkdr["ismonthly"].ToString().Equals("Y")) { NYcheck = true; _startMonth = checkdr["startMonth"].ToString(); _month_frequency = checkdr["month_frequency"].ToString(); }
                }

                //============================================================================================================
                //error setting
                //============================================================================================================
                if (NYcheck == true && YNcheck == true)
                {
                    if (!string.IsNullOrEmpty(dr["sheet_categoryid"].ToString()))
                        abnormal_queryStr = string.IsNullOrEmpty(abnormal_queryStr) ? string.Format(@" (sheet_categoryid='{0}' )", dr["sheet_categoryid"].ToString()) : string.Format(@"{0} Or sheet_categoryid='{1}')", abnormal_queryStr.Replace(")", ""), dr["sheet_categoryid"].ToString());
                }
                //============================================================================================================
                //dayly item(NN)
                //============================================================================================================
                else if (NYcheck == false && YNcheck == false && NNcheck == true)
                {
                    if (!string.IsNullOrEmpty(dr["sheetid"].ToString()))
                        queryStr = string.IsNullOrEmpty(queryStr) ? string.Format(@" (sheetid='{0}' )", dr["sheetid"].ToString()) : string.Format(@"{0} Or sheetid='{1}')", queryStr.Replace(")", ""), dr["sheetid"].ToString());
                }
                //============================================================================================================
                //weekly item(YN)
                //============================================================================================================
                else if (YNcheck == true && NYcheck == false)
                {
                    int nowWeekID = Convert.ToInt32(dateTable.Rows[0]["weekid"].ToString().Substring(2, 2));
                    int startWeekID = Convert.ToInt32(_startWeek);
                    int frequency = Convert.ToInt32(_frequency);

                    if ((nowWeekID - startWeekID) % frequency == 0) //if no remainder, show items
                    {
                        if (!string.IsNullOrEmpty(dr["sheetid"].ToString()))
                            queryStr = string.IsNullOrEmpty(queryStr) ? string.Format(@" (sheetid='{0}' )", dr["sheetid"].ToString()) : string.Format(@"{0} Or sheetid='{1}')", queryStr.Replace(")", ""), dr["sheetid"].ToString());

                    }
                }
                //============================================================================================================
                //monthly item(NY) 
                //============================================================================================================
                else if (YNcheck == false && NYcheck == true)
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

                    //if no remainder, show items(check first week in this month?)
                    if (monthdt.Rows.Count > 0)
                    {
                        int nowMonthID = Convert.ToInt32(monthdt.Rows[0]["month_id"].ToString());
                        int startMonthID = Convert.ToInt32(_startMonth);
                        int Month_frequency = Convert.ToInt32(_month_frequency);

                        //if no remainder, show items(check first week in this month?)
                        if ((nowMonthID - startMonthID) % Month_frequency == 0)
                        {
                            if (!string.IsNullOrEmpty(dr["sheetid"].ToString()))
                                queryStr = string.IsNullOrEmpty(queryStr) ? string.Format(@" (sheetid='{0}' )", dr["sheetid"].ToString()) : string.Format(@"{0} Or sheetid='{1}')", queryStr.Replace(")", ""), dr["sheetid"].ToString());
                        }
                    }
                }//end if

            }//end foreach
            #endregion

            //abnormal
            string ab_Str = string.Format(@"select sheet_categoryid,describes,docnumber
                                            from sheet_category 
                                            where (isdelete is null or isdelete='N') {0}
                                            ", string.IsNullOrEmpty(abnormal_queryStr) ? string.Empty : string.Format(@" And {0}", abnormal_queryStr), abnormal_queryStr);
            System.Data.DataTable ab_logTable = new System.Data.DataTable();
            ab_logTable = ado.loadDataTable(ab_Str, null, "sheet_category");


            //log
            string logStr = string.Format(@"Select Tester,Location,isSealing
                                            From ACS_Manage
                                            Where Tester not in (Select machine
                                                                 From vw_insert_delete_log 
                                                                 Where weekid = :now_weekid
                                                                 And Log_Action ='INSERT'
                                                                 {0}) 
                                                  And Tester in (Select machine 
                                                                 From vw_insert_delete_log
                                                                 Where {1} group by machine)
                                            ", string.IsNullOrEmpty(queryStr) ? string.Empty : string.Format(@" And {0}", queryStr), queryStr);
            para = new object[] { dateTable.Rows[0]["weekid"].ToString() };
            System.Data.DataTable logTable = new System.Data.DataTable();
            logTable = ado.loadDataTable(logStr, para, "ACS_Manage");

            //=================================================================================
            //Excel
            //=================================================================================

            string title = "None Record Machines";
            string title_2 = "Error Setting List";
            HSSFWorkbook hssfWorkBook_1 = new HSSFWorkbook();

            #region sheet1
            HSSFSheet sheet = (NPOI.HSSF.UserModel.HSSFSheet)hssfWorkBook_1.CreateSheet(title);
            IRow row; ICell cell;
            bool isOnly = false;
            int Row_Count = 0;
            int Cell_Count = 0;


            row = sheet.CreateRow(Row_Count);
            string[] xls_title = new string[] { "Tester", "Location", "isSealing", "Process", "Model" };
            short[] colorLst = new short[] { NPOI.HSSF.Util.HSSFColor.YELLOW.index2, NPOI.HSSF.Util.HSSFColor.TEAL.index2, NPOI.HSSF.Util.HSSFColor.SKY_BLUE.index, NPOI.HSSF.Util.HSSFColor.ROSE.index, NPOI.HSSF.Util.HSSFColor.ORANGE.index };
            foreach (string _title in xls_title)
            {
                cell = row.CreateCell(Cell_Count);
                cell.SetCellValue(_title.ToString());
                cell.CellStyle.FillBackgroundColor = colorLst[Cell_Count];
                Cell_Count++;
            }
            if (isOnly == false)
            {
                sheet.SetAutoFilter(CellRangeAddress.ValueOf(string.Format(@"A{0}:E{0}", Row_Count + 1)));//Fliter
                sheet.CreateFreezePane(Cell_Count, Row_Count + 1);//Freeze
                isOnly = true;
            }
            Row_Count += 1;

            //set value
            foreach (DataRow dr in logTable.Rows)
            {
                Cell_Count = 0;
                row = sheet.CreateRow(Row_Count);

                string acsConn = ePM_weekly_Scan.Properties.Settings.Default.ACS;
                Common.AdoDbConn acsAdo = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, acsConn);

                string acsStr = string.Format(@"Select * 
                                                From VW_Machine_List_ALL
                                                Where Machine='{0}'", dr["Tester"].ToString().ToUpper());
                System.Data.DataTable acsDt = acsAdo.loadDataTable(acsStr, null, "VW_Machine_List_ALL");

                //Tester
                cell = row.CreateCell(Cell_Count);
                cell.SetCellValue(dr["Tester"].ToString());
                Cell_Count++;

                //Location
                cell = row.CreateCell(Cell_Count);
                cell.SetCellValue(dr["Location"].ToString());
                Cell_Count++;

                //isSealing
                cell = row.CreateCell(Cell_Count);
                cell.SetCellValue(dr["isSealing"].ToString());
                Cell_Count++;

                if (acsDt.Rows.Count > 0)
                {
                    //Process
                    cell = row.CreateCell(Cell_Count);
                    cell.SetCellValue(acsDt.Rows[0]["Process"].ToString());
                    Cell_Count++;

                    //Model
                    cell = row.CreateCell(Cell_Count);
                    cell.SetCellValue(acsDt.Rows[0]["Model"].ToString());
                    Cell_Count++;
                }

                if (dr["isSealing"].ToString().Equals("Y"))// color red
                {
                    //myRange = (Range)mySheet.get_Range(myExcel.Cells[2 + i, 1], myExcel.Cells[2 + i, 3]);
                    //myRange.Select();
                    //myRange.Columns.Cells.Interior.Color = "255";
                    //myRange.Columns.Cells.Borders.ColorIndex = "0";
                }

                Row_Count++;
            }
            #endregion
            #region sheet2
            HSSFSheet sheet_2 = (NPOI.HSSF.UserModel.HSSFSheet)hssfWorkBook_1.CreateSheet(title_2);
            IRow row_2; ICell cell_2;
            bool isOnly_2 = false;
            int Row_Count_2 = 0;
            int Cell_Count_2 = 0;

            row_2 = sheet_2.CreateRow(Row_Count_2);
            string[] xls_title_2 = new string[] { "sheet_categoryID", "Category_Name", "docnumber" };
            short[] colorLst_2 = new short[] { NPOI.HSSF.Util.HSSFColor.YELLOW.index2, NPOI.HSSF.Util.HSSFColor.TEAL.index2, NPOI.HSSF.Util.HSSFColor.SKY_BLUE.index };
            foreach (string _title in xls_title_2)
            {
                cell_2 = row_2.CreateCell(Cell_Count_2);
                cell_2.SetCellValue(_title.ToString());
                cell_2.CellStyle.FillBackgroundColor = colorLst_2[Cell_Count_2];
                Cell_Count_2++;
            }
            if (isOnly == false)
            {
                sheet_2.SetAutoFilter(CellRangeAddress.ValueOf(string.Format(@"A{0}:C{0}", Row_Count_2 + 1)));//Fliter
                sheet_2.CreateFreezePane(Cell_Count_2, Row_Count_2 + 1);//Freeze
                isOnly_2 = true;
            }
            Row_Count_2 += 1;

            //set value
            foreach (DataRow dr in ab_logTable.Rows)
            {
                Cell_Count_2 = 0;
                row_2 = sheet_2.CreateRow(Row_Count_2);

                //sheet_categoryid
                cell_2 = row_2.CreateCell(Cell_Count_2);
                cell_2.SetCellValue(dr["sheet_categoryid"].ToString());
                Cell_Count_2++;

                //describes
                cell_2 = row_2.CreateCell(Cell_Count_2);
                cell_2.SetCellValue(dr["describes"].ToString());
                Cell_Count_2++;

                //docnumber
                cell_2 = row_2.CreateCell(Cell_Count_2);
                cell_2.SetCellValue(dr["docnumber"].ToString());
                Cell_Count_2++;

                Row_Count_2++;
            }
            #endregion

            //Save File (\download\..)
            string FileName = string.Format(@"{0}_ePM__NoneRecordTester.xls", DateTime.Now.ToString("yyyy-MM-dd"));
            string dir = ePM_weekly_Scan.Properties.Settings.Default.websiteDir;
            string PathFile = dir + FileName;
            FileStream fs = new FileStream(PathFile, FileMode.Create);
            hssfWorkBook_1.Write(fs);
            fs.Close();
            fs.Dispose();


            //=================================================================================
            //Mail
            //=================================================================================            
            if (logTable.Rows.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                string webUrl = ePM_weekly_Scan.Properties.Settings.Default.webSiteUrl;
                string downloadUrl = webUrl + FileName;

                sb.Insert(0, string.Format(@"<p><b>Week {0}</b></p>", dateTable.Rows[0]["weekid"].ToString().Substring(1, 3)));
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
                                                       "Alan.Kuo@nxp.com", "t.w.lin@nxp.com"
                                                     };
                foreach (string contact in contact_List)
                {
                    //func.mail(contact, sb.ToString());
                }
            }
        } //old
    }
}
