using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace RTMReportsAssistant
{
    public class clsInvoice_OIR
    {
        SqlConnection con = new SqlConnection(@"Data Source=10.55.5.40,1433;Initial Catalog=RTM_Global_Test;User ID=PRODRTMDB;Password=Prodrtm@123;");
        DataTable dt = new DataTable();
        string query;
        SqlDataAdapter da;

        public DataTable GetOIRTaskDetails(int teamId, string date)
        {
            dt = new DataTable();

            query = "SELECT CONVERT(VARCHAR(12), R_TimeDate, 101) as [Date], R_User_Name as [User], "+
                    "CL_ClientName as [Client], CL_Code as [Client Code], TL_Task as [Task] ,STL_SubTask as [Sub Task], "+
                    "COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600),3),0) as Duration, "+
                    "R_Comments from RTM_Records WITH (NOLOCK) "+
                    "left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID "+
                    "left join RTM_SubTask_List WITH (NOLOCK) on R_SubTask = STL_ID "+ 
                    "left join RTM_Task_List WITH (NOLOCK) on R_Task = TL_ID "+
                    "where R_TeamId="+ teamId +" "+
                    "and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) = '"+ date +"' "+
                    "and R_Duration != 'HH:MM:SS'  "+
                    "and TL_Task ='OIR' and STL_SubTask ='OIR Invoice Balancing' " +
                    "GROUP BY CONVERT(VARCHAR(12), R_TimeDate, 101), R_User_Name,CL_ClientName, TL_Task, STL_SubTask, R_Comments, CL_Code "+
                    "Order By R_User_Name";

            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dt);
            }

            return dt;
        }
    }
}
