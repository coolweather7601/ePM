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

namespace EPM_Web.Alan
{
    public partial class UpdateLog : System.Web.UI.Page
    {
        public static string Conn = System.Configuration.ConfigurationManager.ConnectionStrings["EPM"].ToString();
        public Common.AdoDbConn ado;
        static public string sheetID;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                sheetID = Request.QueryString["sheetID"];
                initial();
            }
        }
        private void initial()
        {
            ado = new Common.AdoDbConn(Common.AdoDbConn.AdoDbType.Oracle, Conn);
            string strsql = string.Format(@"select Update_logID,item_desc,oldValue,newValue,UpdateTime,account
                                            FROM vw_Update_Log
                                            WHERE SheetID='{0}'
                                            Order by UpdateTime DESC", sheetID);
            DataTable dt = ado.loadDataTable(strsql, null, "vw_Update_Log");
            GridView1.DataSource = dt;
            GridView1.DataBind();
        }

    }
}