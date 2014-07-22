using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Data.OracleClient;
using System.Data.Common;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Collections.Generic;

/// <summary>
/// AdoDbConn 的摘要描述
/// </summary>
namespace EPM.Alan.Common
{

    public class AdoDbConn
    {
        #region GetProviderFactory(...)
        // 不是.net已经自带的DataProvider,必須修改C:\WINDOWS\Microsoft.NET\Framework\v2.0.50727\CONFIG\machine.config
        // 如果程式的執行主機上已有安裝舊版的DataProvider(其他廠商的程式)或沒有權限安裝DataProvider(雲端)
        // 只能用此法; 缺點是使用反射造成的延遲
        private static object GetPublicStaticField(Type type, string fieldName)
        {
            try
            { return type.InvokeMember(fieldName, BindingFlags.Public | BindingFlags.Static | BindingFlags.GetField, null, null, null); }
            catch (Exception e)
            { return null; }
        }
        private static Type LoadType(string assemblyName, string className)
        {
            try
            { return Assembly.Load(assemblyName).GetType(className); }
            catch (Exception e)
            { return null; }
        }

        public static DbProviderFactory GetProviderFactory(string providerName, string assemblyName, string className)
        {
            DbProviderFactory provider;

            try
            {
                provider = DbProviderFactories.GetFactory(providerName);
            }
            catch (Exception)
            {
                // dll檔必須同目錄而且assemblyName不要加副檔名
                provider = GetPublicStaticField(LoadType(assemblyName, className), "Instance") as DbProviderFactory;
            }

            if (provider == null)
            {
                string msg = "Create DbProviderFactory failed.";
                throw new Exception(msg);
            }
            return provider;
        }

        #endregion


        public enum AdoDbType : int
        { MsSql = 0, MsAccess = 1, MySQL = 2, SQLite = 3, Oracle = 4 }

        AdoDbType dbtype = AdoDbType.MsSql;
        bool blnDbConnRegisted = false;
        DbCommand dbCmd = null;
        DbDataAdapter dbDa = null;
        DbCommandBuilder dbCb = null;
        public DataSet dsTemp = null;
        public static Regex theReg = new Regex(@"([@][a-z|A-Z|u4e00-u9fa5]+)");//mssql

        private void Create(bool _blnDbConnRegisted, AdoDbType _AdoDbType, string _AdoConnStr)
        {
            dbtype = _AdoDbType;
            try
            {
                DbProviderFactory dbfc = null;

                blnDbConnRegisted = _blnDbConnRegisted;
                if (_AdoDbType == AdoDbType.MsSql || _AdoDbType == AdoDbType.MsAccess)
                { blnDbConnRegisted = true; } //已內建在system.data.dll 中

                switch (_AdoDbType)
                {
                    case AdoDbType.MsSql: { dbfc = GetProviderFactory("System.Data.SqlClient", null, null); } break;
                    case AdoDbType.MsAccess: { dbfc = GetProviderFactory("System.Data.OleDb", null, null); } break;

                    case AdoDbType.MySQL:
                        {
                            if (blnDbConnRegisted) { dbfc = GetProviderFactory("MySql.Data.MySqlClient", null, null); }
                            else { dbfc = GetProviderFactory(null, "MySql.Data", "MySql.Data.MySqlClient.MySqlClientFactory"); }
                        } break;

                    case AdoDbType.SQLite:
                        {
                            if (blnDbConnRegisted) { dbfc = GetProviderFactory("System.Data.SQLite", null, null); }
                            else { dbfc = GetProviderFactory(null, "System.Data.SQLite", "System.Data.SQLite.SQLiteFactory"); }
                        } break;
                    case AdoDbType.Oracle:
                        {
                            dbfc = GetProviderFactory("System.Data.OracleClient", null, null);
                            theReg = new Regex(@"([:][a-z|A-Z|u4e00-u9fa5]+)");
                        } break;
                    default: break;
                }

                if (dbfc != null)
                {
                    dbCmd = dbfc.CreateCommand();
                    dbCmd.Connection = dbfc.CreateConnection();
                    dbCmd.Connection.ConnectionString = _AdoConnStr;

                    dbDa = dbfc.CreateDataAdapter();
                    dbDa.SelectCommand = dbCmd;

                    dbCb = dbfc.CreateCommandBuilder();
                    dbCb.DataAdapter = dbDa;

                    dsTemp = new DataSet("dsTemp");
                }
            }
            catch (Exception ex)
            { string err = ex.Message; }
        }
        public AdoDbConn(bool _blnDbConnRegisted, AdoDbType _AdoDbType, string _AdoConnStr)
        { Create(_blnDbConnRegisted, _AdoDbType, _AdoConnStr); }
        public AdoDbConn(AdoDbType _AdoDbType, string _AdoConnStr)
        { Create(false, _AdoDbType, _AdoConnStr); }
        public AdoDbConn(int _AdoDbType, string _AdoConnStr)
        { Create(false, (AdoDbType)_AdoDbType, _AdoConnStr); }

        private enum AdoDbAction : int
        { Select = 0, Update = 1, ExecuteNonQuery = 2, InsertExecuteReader = 3 }
        private object accessDataTable(AdoDbAction action, string strSql, object[] sqlParams, string tableNameInDs)
        {
            object re = null;

            try
            {
                dbCmd.Parameters.Clear();

                //Regex theReg = new Regex(@"([:][a-z|A-Z|u4e00-u9fa5]+)");//oracle 
                //Regex theReg = new Regex(@"([@][a-z|A-Z|u4e00-u9fa5]+)");//mssql 

                MatchCollection mc = theReg.Matches(strSql);
                if (sqlParams != null && mc.Count == sqlParams.Length)
                {
                    for (int i = 0; i < mc.Count; i++)
                    {
                        DbParameter dbp = dbCmd.CreateParameter();
                        dbp.ParameterName = mc[i].ToString();
                        dbp.Value = sqlParams[i];
                        dbCmd.Parameters.Add(dbp);
                    }
                }

                dbCmd.CommandText = strSql;
                if (action == AdoDbAction.Select)
                {
                    if (dsTemp.Tables[tableNameInDs] == null)
                    { dsTemp.Tables.Add(tableNameInDs); }
                    else
                    {
                        dsTemp.Tables[tableNameInDs].Dispose();
                        dsTemp.Tables.Remove(tableNameInDs);
                        dsTemp.Tables.Add(tableNameInDs);
                    }
                }

                switch (action)
                {
                    case AdoDbAction.Select: { dbDa.Fill(dsTemp, tableNameInDs); re = dsTemp.Tables[tableNameInDs]; } break;
                    case AdoDbAction.Update: { dbDa.Update(dsTemp, tableNameInDs); } break;
                    case AdoDbAction.ExecuteNonQuery: { dbCmd.Connection.Open(); dbCmd.ExecuteNonQuery(); } break;
                    case AdoDbAction.InsertExecuteReader:
                        {
                            dbCmd.CommandText += ";SELECT NewID = SCOPE_IDENTITY()";
                            dbCmd.Connection.Open();
                            DbDataReader dr = dbCmd.ExecuteReader();
                            if (dr.HasRows)
                            {
                                dr.Read();
                                re = Convert.ToInt32(dr["NewID"]);
                            }
                            dr.Close();
                        } break;
                    default: break;
                }
            }
            catch (Exception ex)
            { string err = ex.Message; }
            finally
            { if (dbCmd != null && dbCmd.Connection != null) dbCmd.Connection.Close(); }
            return re;
        }

        public DataTable loadDataTable(string strSql, object[] sqlParams, string tableNameInDs)
        { return (DataTable)accessDataTable(AdoDbAction.Select, strSql, sqlParams, tableNameInDs); }
        public void updateDataTable(string strSql, object[] sqlParams, string tableNameInDs)
        { accessDataTable(AdoDbAction.Update, strSql, sqlParams, tableNameInDs); }
        public void dbNonQuery(string strSql, object[] sqlParams)
        { accessDataTable(AdoDbAction.ExecuteNonQuery, strSql, sqlParams, null); }
        public int insertAndGetIdentity(string strSql, object[] sqlParams)
        { return (int)accessDataTable(AdoDbAction.InsertExecuteReader, strSql, sqlParams, null); }
        public string SQL_transaction(List<string> arrStr, string Conn)
        {
            string reStr = "SUCCESS";
            using (OracleConnection connection = new OracleConnection(Conn))
            {
                connection.Open();

                OracleCommand command = connection.CreateCommand();
                OracleTransaction transaction;

                // Start a local transaction
                transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);
                // Assign transaction object for a pending local transaction
                command.Transaction = transaction;

                try
                {
                    foreach (string str in arrStr)
                    {
                        command.CommandText = str;
                        command.ExecuteNonQuery();
                    }
                    transaction.Commit();
                    Console.WriteLine("Both records are written to database.");
                }
                catch (Exception e)
                {
                    reStr = string.Format(@"{0}", e.ToString());
                    transaction.Rollback();
                    Console.WriteLine(e.ToString());
                    Console.WriteLine("Neither record was written to database.");
                }
                return reStr;
            }
        }

    }
}
