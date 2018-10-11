using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using RTMReportsAssistant;

namespace ResourceUtil
{
    public partial class Form1 : Form
    {
    //Database connection string.
        SqlConnection con = new SqlConnection(@"Data Source=10.55.5.40,1433;Initial Catalog=RTM_Global_Test;User ID=PRODRTMDB;Password=Prodrtm@123;");

        DataTable dtTeam = new System.Data.DataTable();
        DataTable dt = new System.Data.DataTable();
        DataTable dtResult = new System.Data.DataTable();

        SqlDataAdapter da = new SqlDataAdapter();
        string query;
        public Form1()
        {
            InitializeComponent();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            lblStatus.Text = "Please wait....";
            ExportDetails();
        }

        private DataTable Teams()
        {
            dtTeam = new DataTable();
            //india
          // query = "select T_ID, T_TeamName from RTM_Team_List WITH (NOLOCK) WHERE T_Location = 'IND' ORDER BY T_ID";

            //China
              query = "select T_ID, T_TeamName from RTM_Team_List WITH (NOLOCK) WHERE T_Location = 'CHN' ORDER BY T_ID";

            //only for melissa team***********************
            //query = "select T_ID, T_TeamName from RTM_Team_List WITH (NOLOCK) WHERE T_ID IN (56,57,58,60,61,63,66,69,73,77,99,101,130) ORDER BY T_ID";

            //Romenia
           // query = "select T_ID, T_TeamName from RTM_Team_List WITH (NOLOCK) WHERE T_Id in (133,134,135) ORDER BY T_ID";

            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dtTeam);
            }

            return dtTeam;
        }

        private DataTable GetHours(int teamId, string from, string to)
        {
            dt = new System.Data.DataTable();

            if (rbWithMgr.Checked)
            {
                query = "SELECT " +
                            "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600),2),0) from RTM_Records WITH (NOLOCK) left join RTM_SubTask_List WITH (NOLOCK) on R_SubTask = STL_ID left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where R_TeamId = " + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and R_Duration != 'HH:MM:SS' and CL_ClientName <> 'Internal' and R_Task <> 0) as [Billable Hours], " +
                            "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600),2),0) from RTM_Records where R_TeamId = " + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and R_Duration != 'HH:MM:SS' and R_Task = 0) + " +
                            "(SELECT COALESCE(SUM(CAST((CASE WHEN LD_Duration IS NULL THEN 8 ELSE CAST(LD_Duration as float) END) as float)),0) as [hours] from RTM_LeaveDetails WITH (NOLOCK) left join RTM_User_List ON LD_UserName = UL_User_Name where UL_Team_Id=" + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LD_Date))) BETWEEN '" + from + "' and '" + to + "') as [PTO], " +
                            "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600),2),0) from RTM_Records WITH (NOLOCK) left join RTM_SubTask_List WITH (NOLOCK) on R_SubTask = STL_ID left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where R_TeamId = " + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and R_Duration != 'HH:MM:SS' and STL_Subtask Like '%Employee Engagement%') + " +
                            "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),1) as float))/3600),2),0) as [hours] from RTM_Log_Actions WITH (NOLOCK) where LA_TeamId=" + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and LA_Duration != 'HH:MM:SS' and LA_Reason ='Meeting' AND LA_Comments LIKE '%Employee Engagement%') as [EEA], " +
                            "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600),2),0) from RTM_Records WITH (NOLOCK) left join RTM_SubTask_List WITH (NOLOCK) on R_SubTask = STL_ID left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where R_TeamId = " + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and R_Duration != 'HH:MM:SS' and CL_ClientName = 'Internal' and STL_Subtask = 'NON-TASK') + " +
                            "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),1) as float))/3600),2),0) as [hours] from RTM_Log_Actions WITH (NOLOCK) where LA_TeamId=" + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and LA_Duration != 'HH:MM:SS' and (LA_Reason ='Non-Task')) as [Nontask], " +
                            "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600),2),0) from RTM_Records WITH (NOLOCK) left join RTM_SubTask_List WITH (NOLOCK) on R_SubTask = STL_ID left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where R_TeamId = " + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and R_Duration != 'HH:MM:SS' and CL_ClientName = 'Internal' and (STL_Subtask Like '%Meeting%' or STL_Subtask Like '%Training%' or STL_SubTask Like '%Learning%')) + " +
                            "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),1) as float))/3600),2),0) as [hours] from RTM_Log_Actions WITH (NOLOCK) where LA_TeamId=" + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and LA_Duration != 'HH:MM:SS' and LA_Reason ='Meeting' AND LA_Comments NOT LIKE '%Employee Engagement%') as [Meeting], " +
                            "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600),2),0) from RTM_Records WITH (NOLOCK) left join RTM_SubTask_List WITH (NOLOCK) on R_SubTask = STL_ID left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where R_TeamId = " + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and R_Duration != 'HH:MM:SS' and CL_ClientName = 'Internal' and (STL_Subtask NOT Like '%Meeting%' AND STL_Subtask NOT Like '%Training%' AND STL_SubTask NOT Like '%Learning%' AND STL_Subtask <> 'NON-TASK' AND STL_Subtask <> 'Available' AND STL_Subtask NOT Like '%Employee Engagement%')) + " +
                            "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),1) as float))/3600),2),0) as [hours] from RTM_Log_Actions WITH (NOLOCK) where LA_TeamId=" + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and LA_Duration != 'HH:MM:SS' and (LA_Reason = 'Conference-Call' or LA_Reason = 'Conf-Call' or LA_Reason ='Peer Support')) as [Others]," +
                            "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600),2),0) from RTM_Records WITH (NOLOCK) left join RTM_SubTask_List WITH (NOLOCK) on R_SubTask = STL_ID left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where R_TeamId = " + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and R_Duration != 'HH:MM:SS' and CL_ClientName = 'Internal' and (STL_Subtask = 'Available')) as [Available Time]";
            }
            else
            {
                query = "SELECT " +
                        "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600),2),0) from RTM_Records WITH (NOLOCK) left join RTM_SubTask_List WITH (NOLOCK) on R_SubTask = STL_ID left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where R_TeamId = " + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and R_Duration != 'HH:MM:SS' and CL_ClientName <> 'Internal' and R_Task <> 0 " +
                         "AND R_User_Name NOT IN (select LeadName from tblLeads WITH (NOLOCK))) as [Billable Hours], " +
                        "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600),2),0) from RTM_Records where R_TeamId = " + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and R_Duration != 'HH:MM:SS' and R_Task = 0 " +
                         "AND R_User_Name NOT IN (select LeadName from tblLeads WITH (NOLOCK))) + " +
                        "(SELECT COALESCE(SUM(CAST((CASE WHEN LD_Duration IS NULL THEN 8 ELSE LD_Duration END) as float)),0) as [hours] from RTM_LeaveDetails WITH (NOLOCK) left join RTM_User_List ON LD_UserName = UL_User_Name where UL_Team_Id=" + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LD_Date))) BETWEEN '" + from + "' and '" + to + "' " +
                        "AND LD_UserName NOT IN (select LeadName from tblLeads WITH (NOLOCK))) as [PTO], " +
                        "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600),2),0) from RTM_Records WITH (NOLOCK) left join RTM_SubTask_List WITH (NOLOCK) on R_SubTask = STL_ID left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where R_TeamId = " + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and R_Duration != 'HH:MM:SS' and STL_Subtask Like '%Employee Engagement%' " +
                        " AND R_User_Name NOT IN(select LeadName from tblLeads WITH (NOLOCK))) + " +
                        "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),1) as float))/3600),2),0) as [hours] from RTM_Log_Actions WITH (NOLOCK) where LA_TeamId=" + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and LA_Duration != 'HH:MM:SS' and LA_Reason ='Meeting' AND LA_Comments LIKE '%Employee Engagement%' " +
                        " AND LA_User_Name NOT IN (select LeadName from tblLeads WITH (NOLOCK))) as [EEA], " +
                        "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600),2),0) from RTM_Records WITH (NOLOCK) left join RTM_SubTask_List WITH (NOLOCK) on R_SubTask = STL_ID left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where R_TeamId = " + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and R_Duration != 'HH:MM:SS' and CL_ClientName = 'Internal' and STL_Subtask = 'NON-TASK' " +
                        " AND R_User_Name NOT IN (select LeadName from tblLeads WITH (NOLOCK))) + " +
                        "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),1) as float))/3600),2),0) as [hours] from RTM_Log_Actions WITH (NOLOCK) where LA_TeamId=" + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and LA_Duration != 'HH:MM:SS' and (LA_Reason ='Non-Task') " +
                        " AND LA_User_Name NOT IN (select LeadName from tblLeads WITH (NOLOCK))) as [Nontask], " +
                        "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600),2),0) from RTM_Records WITH (NOLOCK) left join RTM_SubTask_List WITH (NOLOCK) on R_SubTask = STL_ID left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where R_TeamId = " + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and R_Duration != 'HH:MM:SS' and CL_ClientName = 'Internal' and (STL_Subtask Like '%Meeting%' or STL_Subtask Like '%Training%' or STL_SubTask Like '%Learning%') " +
                        " AND R_User_Name NOT IN (select LeadName from tblLeads WITH (NOLOCK))) + " +
                        "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),1) as float))/3600),2),0) as [hours] from RTM_Log_Actions WITH (NOLOCK) where LA_TeamId=" + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and LA_Duration != 'HH:MM:SS' and LA_Reason ='Meeting' AND LA_Comments NOT LIKE '%Employee Engagement%'  " +
                        " AND LA_User_Name NOT IN (select LeadName from tblLeads WITH (NOLOCK))) as [Meeting], " +
                        "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600),2),0) from RTM_Records WITH (NOLOCK) left join RTM_SubTask_List WITH (NOLOCK) on R_SubTask = STL_ID left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where R_TeamId = " + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and R_Duration != 'HH:MM:SS' and CL_ClientName = 'Internal' and (STL_Subtask NOT Like '%Meeting%' AND STL_Subtask NOT Like '%Training%' AND STL_SubTask NOT Like '%Learning%' and STL_Subtask <> 'NON-TASK' AND STL_Subtask <> 'Available' AND STL_Subtask NOT Like '%Employee Engagement%') " +
                        " AND R_User_Name NOT IN (select LeadName from tblLeads WITH (NOLOCK))) + " +
                        "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(LA_Duration,':','.'),1) as float))/3600),2),0) as [hours] from RTM_Log_Actions WITH (NOLOCK) where LA_TeamId=" + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and LA_Duration != 'HH:MM:SS' and (LA_Reason = 'Conference-Call' or LA_Reason = 'Conf-Call' or LA_Reason ='Peer Support') " +
                        " AND LA_User_Name NOT IN (select LeadName from tblLeads WITH (NOLOCK))) as [Others], " +
                        "(SELECT COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600),2),0) from RTM_Records WITH (NOLOCK) left join RTM_SubTask_List WITH (NOLOCK) on R_SubTask = STL_ID left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where R_TeamId = " + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) BETWEEN '" + from + "' and '" + to + "' and R_Duration != 'HH:MM:SS' and CL_ClientName = 'Internal' and (STL_Subtask = 'Available') " +
                        " AND R_User_Name NOT IN (select LeadName from tblLeads WITH (NOLOCK))) as [Available Time]";
            }
            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dt);
            }

            return dt;
        }

        private DataTable GetUsersCount(int teamId, string from, string to)
        {
            dt = new System.Data.DataTable();
            if (rbWithMgr.Checked)
            {
                query = "select count(Distinct R_User_Name) [UsersCount] from RTM_Records WHERE R_TeamId = " + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) BETWEEN '" + from + "' and '" + to + "'";
            }
            else
            {
                query = "select count(Distinct R_User_Name) [UsersCount] from RTM_Records WHERE R_TeamId = " + teamId + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) BETWEEN '" + from + "' and '" + to + "' AND R_User_Name NOT IN (select LeadName from tblLeads WITH (NOLOCK))";
            }
            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dt);
            }

            return dt;
        }

        public int GetNumberOfWorkingDays(DateTime start, DateTime stop)
        {
            TimeSpan interval = stop - start;
            int totalWorkingDays;
            checked
            {
                int totalWeek = interval.Days / 7;
                totalWorkingDays = 5 * totalWeek;

                int remainingDays = interval.Days % 7;


                for (int i = 0; i <= remainingDays; i++)
                {
                    DayOfWeek test = (DayOfWeek)(((int)start.DayOfWeek + i) % 7);
                    if (test >= DayOfWeek.Monday && test <= DayOfWeek.Friday)
                        totalWorkingDays++;
                }
            }

            return totalWorkingDays;
        }

        private void BuildTable()
        {
            dtResult = new System.Data.DataTable();
            dtResult.Columns.Add("TeamName");
            dtResult.Columns.Add("DateRange");
            dtResult.Columns.Add("BillableHours");
            dtResult.Columns.Add("PTOs");
            dtResult.Columns.Add("EEA");
            dtResult.Columns.Add("Meeting");
            dtResult.Columns.Add("Nontask");
            dtResult.Columns.Add("Others");
            dtResult.Columns.Add("Available Time");
            dtResult.Columns.Add("UserCount");
            dtResult.Columns.Add("WorkingHours");
        }

        private void ExportDetails()
        {
            try
            {

                DataRow dr;
                BuildTable();
                string filePath = "";
                string fromDate = dpFrom.Value.ToShortDateString();
                string toDate = dpTo.Value.ToShortDateString();
                dtTeam = Teams();

                for (int i = 0; i <= dtTeam.Rows.Count - 1; i++)
                {
                    dr = dtResult.NewRow();
                    int teamId = Convert.ToInt32(dtTeam.Rows[i]["T_ID"]);

                    string teamName = dtTeam.Rows[i]["T_TeamName"].ToString();
                    dr["TeamName"] = teamName;



                    dr["DateRange"] = fromDate + " - " + toDate;

                    int totalWorkingDays = GetNumberOfWorkingDays(dpFrom.Value, dpTo.Value);

                    dt = new System.Data.DataTable();

                    dt = GetHours(teamId, fromDate, toDate);

                    if (dt.Rows.Count > 0)
                    {
                        dr["BillableHours"] = dt.Rows[0]["Billable Hours"];
                        dr["PTOs"] = dt.Rows[0]["PTO"];
                        dr["EEA"] = dt.Rows[0]["EEA"];
                        dr["Meeting"] = dt.Rows[0]["Meeting"];
                        dr["Nontask"] = dt.Rows[0]["Nontask"];
                        dr["Others"] = dt.Rows[0]["Others"];
                        dr["Available Time"] = dt.Rows[0]["Available Time"];
                    }

                    dr["WorkingHours"] = totalWorkingDays * 8;

                    dt = new System.Data.DataTable();

                    dt = GetUsersCount(teamId, fromDate, toDate);
                    if (dt.Rows.Count > 0)
                    {
                        dr["UserCount"] = dt.Rows[0]["UsersCount"];
                    }

                    dtResult.Rows.Add(dr);
                }

                if (rbWithMgr.Checked)
                {
                    filePath = "D:\\Util\\" + dpFrom.Value.ToString("MM-dd-yyyy") + "-" + dpTo.Value.ToString("MM-dd-yyyy") + ".csv";
                }
                else
                {
                    filePath = "D:\\Util\\" + dpFrom.Value.ToString("MM-dd-yyyy") + "-" + dpTo.Value.ToString("MM-dd-yyyy") + "(WithoutManager).csv";
                }

                CSVUtility.ToCSV(dtResult, filePath);

                MessageBox.Show("RU Report Generated");

            }
            catch (Exception ex)
            {
                lblStatus.Text = ex.Message;
            }
        }
    }
}
