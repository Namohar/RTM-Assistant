using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Web.Services;

namespace WebApplication1
{
    public partial class tree : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        public DataTable getDataTable()
        {
            SqlConnection con = new SqlConnection(@"Data Source=BLRPRODRTM\RTM_PROD_BLR;Initial Catalog=Real_Time_Metrics;User ID=sa;Password=Prodrtm@123;");
            DataTable dt = new DataTable();
            string query = " select a.UL_User_Name as EmpName,a.UL_Employee_Id as empID,b.UL_User_Name as MgrName,b.UL_Employee_Id as mgrID ";
            query += " from RTM_User_List a inner join RTM_User_List b on a.UL_RepMgrId=b.UL_Employee_Id";
            SqlDataAdapter dap = new SqlDataAdapter(query, con);
            DataSet ds = new DataSet();
            dap.Fill(ds);
            return ds.Tables[0];
        }

        [WebMethod]
        public List<Google_org_data> getOrgData()
        {
            List<Google_org_data> g = new List<Google_org_data>();
            DataTable myData = getDataTable();

            g.Add(new Google_org_data
            {
                Employee = "Rashmi Ahuja",
                Manager = "",
                mgrID = "",
                empID = "102651",
               
            });

            foreach (DataRow row in myData.Rows)
            {
                string empName = row["EmpName"].ToString();
                var mgrName = row["MgrName"].ToString();
                var mgrID = row["mgrID"].ToString();
                var empID = row["empID"].ToString();
                

                g.Add(new Google_org_data
                {
                    Employee = empName,
                    Manager = mgrName,
                    mgrID = mgrID,
                    empID = empID,
                });
            }
            return g;
        }
    }
}