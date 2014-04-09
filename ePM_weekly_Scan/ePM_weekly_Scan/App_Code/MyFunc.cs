using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;
using System.Data;

namespace EPM.Alan.Common
{
    class MyFunc
    {
        public void mail(string contact, string mail_data)
        {
            string title = "EPM";
            using (OracleConnection connection = new OracleConnection(ePM_weekly_Scan.Properties.Settings.Default.Mail))
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
                    "values (" + mail_id + ",'eStartupCheck mail agent',0,sysdate,0,'" + contact + "',null,'" + title + "',sysdate+1,0,null,1)";
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

    }
}
