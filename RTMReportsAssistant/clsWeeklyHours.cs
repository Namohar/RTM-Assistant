using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace RTMReportsAssistant
{
    public class clsWeeklyHours
    {
        SqlConnection con = new SqlConnection(@"Data Source=10.55.5.40,1433;Initial Catalog=RTM_Global_Test;User ID=PRODRTMDB;Password=Prodrtm@123;");
        DataTable dt = new DataTable();
        string query;
        SqlDataAdapter da;

        public DataTable GetWeeklyHoursLessThan38(string from, string to, string location)
        {
            dt = new DataTable();
            query = "SELECT ROW_NUMBER() over (order by A.UL_RepMgrEmail) as [Sl.No.], A.UL_Employee_ID AS [Employee ID], A.username as [Employee Name], A.team as [Team Name],   A.UL_RepMgrEmail as [Reporting Manager], ROUND(SUM(COALESCE(A.Sum_Col1,0) + COALESCE(B.Sum_Col1,0)),2) as [Total Hours] " +
                    "FROM ( select T_TeamName as [team], R_User_Name as [username], UL_Employee_ID, UL_RepMgrEmail, SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600) as Sum_Col1 FROM  RTM_Records WITH (NOLOCK) left join RTM_Team_List WITH (NOLOCK) on R_TeamId = T_ID left join RTM_User_List WITH (NOLOCK) on R_User_Name= UL_User_Name where R_Duration !='HH:MM:SS' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) Between '"+ from +"' and '"+ to +"' and T_Location ='"+ location +"' Group by T_TeamName, R_User_Name, UL_Employee_ID, UL_RepMgrEmail  ) A  "+
                    "Left join ( select T_TeamName as [team], LA_User_Name as [username], UL_Employee_ID, UL_RepMgrEmail, SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),1) as float))/3600) as Sum_Col1 FROM  RTM_Log_Actions WITH (NOLOCK) left join RTM_Team_List WITH (NOLOCK) on LA_TeamId = T_ID left join RTM_User_List WITH (NOLOCK) on LA_User_Name = UL_User_Name where LA_Duration !='HH:MM:SS' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_TimeDate))) Between '" + from + "' and '" + to + "' and T_Location ='"+ location +"' and (LA_Reason = 'Non-Task' or LA_Reason = 'Conference-Call' or LA_Reason = 'Conf-Call' or LA_Reason='Meeting' or LA_Reason= 'Meetings' ) Group by T_TeamName, LA_User_Name, UL_Employee_ID, UL_RepMgrEmail) B on A.username = B.username " +
                    "Group by A.team, B.team,A.username, B.username,A.UL_Employee_ID, A.UL_RepMgrEmail HAVING ROUND(SUM(COALESCE(A.Sum_Col1,0) + COALESCE(B.Sum_Col1,0)), 2) <= 38.75 order by A.UL_RepMgrEmail";

            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dt);
            }

            return dt;
        }

        public DataTable GetWeeklyHoursGreaterThan45(string from, string to, string location)
        {
            dt = new DataTable();
            query = "SELECT ROW_NUMBER() over (order by A.UL_RepMgrEmail) as [Sl.No.], A.UL_Employee_ID AS [Employee ID], A.username as [Employee Name], A.team as [Team Name],   A.UL_RepMgrEmail as [Reporting Manager], ROUND(SUM(COALESCE(A.Sum_Col1,0) + COALESCE(B.Sum_Col1,0)),2) as [Total Hours] " +
                    "FROM ( select T_TeamName as [team], R_User_Name as [username], UL_Employee_ID, UL_RepMgrEmail, SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600) as Sum_Col1 FROM  RTM_Records WITH (NOLOCK) left join RTM_Team_List WITH (NOLOCK) on R_TeamId = T_ID left join RTM_User_List WITH (NOLOCK) on R_User_Name= UL_User_Name where R_Duration !='HH:MM:SS' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) Between '" + from + "' and '" + to + "' and T_Location ='" + location + "' Group by T_TeamName, R_User_Name, UL_Employee_ID, UL_RepMgrEmail  ) A  " +
                    "Left join ( select T_TeamName as [team], LA_User_Name as [username], UL_Employee_ID, UL_RepMgrEmail, SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),1) as float))/3600) as Sum_Col1 FROM  RTM_Log_Actions WITH (NOLOCK) left join RTM_Team_List WITH (NOLOCK) on LA_TeamId = T_ID left join RTM_User_List WITH (NOLOCK) on LA_User_Name = UL_User_Name where LA_Duration !='HH:MM:SS' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_TimeDate))) Between '" + from + "' and '" + to + "' and T_Location ='" + location + "' and (LA_Reason = 'Non-Task' or LA_Reason = 'Conference-Call' or LA_Reason = 'Conf-Call' or LA_Reason='Meeting' or LA_Reason= 'Meetings' ) Group by T_TeamName, LA_User_Name, UL_Employee_ID, UL_RepMgrEmail) B on A.username = B.username " +
                    "Group by A.team, B.team,A.username, B.username,A.UL_Employee_ID, A.UL_RepMgrEmail HAVING ROUND(SUM(COALESCE(A.Sum_Col1,0) + COALESCE(B.Sum_Col1,0)), 2) >= 45 order by A.UL_RepMgrEmail";

            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dt);
            }

            return dt;
        }
    }
}
