using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;


namespace RTMReportsAssistant
{
    public class RPAReport
    {
        SqlConnection con = new SqlConnection(@"Data Source=10.55.5.40,1433;Initial Catalog=RTM_Global_Test;User ID=PRODRTMDB;Password=Prodrtm@123;");
        DataTable dt = new DataTable();
        string query;
        SqlDataAdapter da;

        public DataTable GetRPSTaskDetails(string from, string to)
        {
            dt = new DataTable();

            query = "select Convert(Varchar(10), R_TimeDate, 101) as [Date], T_TeamName as [Team], R_User_Name as [User], CL_ClientName as [Client], TL_Task as [Task], "+
                    "STL_SubTask as [SubTask], ROUND(SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600), 2) as [Duration] "+
                    "from RTM_Records "+
                    "left join RTM_Team_List on R_TeamId = T_ID "+
                    "left join RTM_Client_List on R_Client = CL_ID "+
                    "left join RTM_Task_List on R_Task=TL_ID "+
                    "left join RTM_SubTask_List on R_SubTask = STL_ID "+
                    "WHERE TL_Task = 'RPA' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) between '"+ from +"' and '"+ to +"' "+
                    "Group by Convert(Varchar(10), R_TimeDate, 101), T_TeamName, R_User_Name, CL_ClientName, TL_Task, STL_SubTask "+
                    "Order by Convert(Varchar(10), R_TimeDate, 101)";

            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dt);
            }

            return dt;
        }
    }
}
