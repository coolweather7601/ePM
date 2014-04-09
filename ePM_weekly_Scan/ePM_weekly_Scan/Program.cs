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

            //log
            string logStr = @"Select Tester,Location,isSealing
                              From ACS_Manage
                              Where Tester not in (Select machine
                                                   From vw_insert_delete_log 
                                                   Where weekid = :now_weekid
                                                         And Log_Action ='INSERT') ";
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

            int i = 0;
            foreach (System.Data.DataRow dr in logTable.Rows)
            {              
                myExcel.Cells[2 + i, 1] = "'" + dr["Tester"].ToString();
                myExcel.Cells[2 + i, 2] = "'" + dr["Location"].ToString();
                myExcel.Cells[2 + i, 3] = "'" + dr["isSealing"].ToString();

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
                sb.Append(string.Format(@"<table>
                                              <tr><td colspan='2'>None Record Tester</td></tr>
                                              <tr>
                                                  <td>Url：</td>
                                                  <td>{0}</td>
                                              </tr>
                                          </table>", downloadUrl));
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
                                                       "sj.chen@nxp.com","abec.ts@nxp.com","abes.ts@nxp.com","wayne.huang@nxp.com","rudy.chen@nxp.com"
                                                     };

                foreach (string contact in contact_List)
                {
                    func.mail(contact, sb.ToString());
                }
            }  
        }
    }
}
