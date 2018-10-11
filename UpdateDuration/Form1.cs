using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Net.Mail;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;

namespace UpdateDuration
{
    public partial class Form1 : Form
    {

        SqlConnection con = new SqlConnection(@"Data Source=BLRPRODRTM\RTM_PROD_BLR;Initial Catalog=Real_Time_Metrics;User ID=sa;Password=Prodrtm@123;");
        DataSet ds = new DataSet();
        SqlDataAdapter da;
        SqlCommand cmd;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
           // DateTime date2 =Convert.ToDateTime("2015-07-20 08:41:31.630");
            //updatePeerTime();
    
           // UpdateData();
        }

        private DataSet FetchIncorrectDuration()
        {
            if (ds.Tables.Contains("incorrect"))
            {
                ds.Tables.Remove(ds.Tables["incorrect"]);
            }
            da = new SqlDataAdapter("select * from RTM_Records where R_Duration LIKE '-%' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_Start_Date_Time)))= '7/10/2015' and R_TeamId=1", con);
            da.Fill(ds, "incorrect");
            return ds;
        }

        private DataSet GetLogHours(string user, DateTime start, DateTime end)
        {
            if (ds.Tables.Contains("Log"))
            {
                ds.Tables.Remove(ds.Tables["Log"]);
            }
            da = new SqlDataAdapter("SELECT sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/3600 as hour, (sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/60)%60 as minute,(sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/60)%60)))%60 as seconds from RTM_Log_Actions where LA_User_Name='" + user + "' and LA_Start_Date_Time BETWEEN '" + start + "' and '" + end + "' and LA_Duration != 'HH:MM:SS'", con);
            da.Fill(ds, "Log");
            return ds;
        }

        private DataSet GetTaskHours(string user, DateTime start, DateTime end)
        {
            if (ds.Tables.Contains("Task"))
            {
                ds.Tables.Remove(ds.Tables["Task"]);
            }
            da = new SqlDataAdapter("SELECT sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600 as hour, (sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/60)%60 as minute,(sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/60)%60)))%60 as seconds from RTM_Records , RTM_SubTask_List where R_SubTask = STL_ID and R_User_Name='" + user + "' and R_Start_Date_Time BETWEEN '" + start.AddSeconds(20) + "' and '" + end.AddSeconds(-20) + "' and R_Duration != 'HH:MM:SS'", con);
            da.Fill(ds, "Task");
            return ds;
        }

        private void UpdateData()
        {
            TimeSpan totalLog = TimeSpan.Parse("00:00:00");
            TimeSpan totalTask = TimeSpan.Parse("00:00:00");
            TimeSpan taskDiff = TimeSpan.Parse("00:00:00");
            TimeSpan totalLogTask = TimeSpan.Parse("00:00:00");
            TimeSpan totalDuration = TimeSpan.Parse("00:00:00");
            try
            {
                ds = FetchIncorrectDuration();

                if (ds.Tables["incorrect"].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds.Tables["incorrect"].Rows)
                    {
                        totalLog = TimeSpan.Parse("00:00:00");
                        totalTask = TimeSpan.Parse("00:00:00");
                        taskDiff = TimeSpan.Parse("00:00:00");
                        totalLogTask = TimeSpan.Parse("00:00:00");
                        totalDuration = TimeSpan.Parse("00:00:00");
                        int id = Convert.ToInt32(dr["R_ID"]);
                        string name = dr["R_User_Name"].ToString();
                        DateTime startTime = Convert.ToDateTime(dr["R_Start_Date_Time"]);
                        DateTime endTime = Convert.ToDateTime(dr["R_CreatedOn"]);
                        taskDiff = TimeSpan.Parse((endTime - startTime).ToString(@"hh\:mm\:ss"));
                        ds = GetLogHours(name, startTime, endTime);

                        if (ds.Tables["Log"].Rows.Count > 0 && ds.Tables["Log"].Rows[0]["hour"].ToString().Length > 0)
                        {
                            totalLog = TimeSpan.Parse(ds.Tables["Log"].Rows[0]["hour"] + ":" + ds.Tables["Log"].Rows[0]["minute"] + ":" + ds.Tables["Log"].Rows[0]["seconds"]);
                        }

                        ds = GetTaskHours(name, startTime, endTime);
                        if (ds.Tables["Task"].Rows.Count > 0 && ds.Tables["Task"].Rows[0]["hour"].ToString().Length>0)
                        {
                            totalTask = TimeSpan.Parse(ds.Tables["Task"].Rows[0]["hour"] + ":" + ds.Tables["Task"].Rows[0]["minute"] + ":" + ds.Tables["Task"].Rows[0]["seconds"]);
                        }

                        totalLogTask = totalLog.Add(totalTask);

                        totalDuration = taskDiff.Subtract(totalLogTask);

                        cmd = new SqlCommand("update RTM_Records set R_Duration='"+ totalDuration +"' where R_ID="+ id +"", con);
                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }
            }
            catch (Exception)
            {
                
               
            }
           
        }

        private DataSet FetchPeerSupportDataFromLogs()
        {
            da = new SqlDataAdapter("SELECT * from RTM_Log_Actions where LA_Reason = 'Peer Support' and LA_Start_Date_Time BETWEEN '7/27/2015' and '8/02/2015' and LA_TeamId='18'", con);
            da.Fill(ds, "LogPeer");
            return ds;
        }

        private DataTable FetchPeerRecord(DateTime start, string user)
        {
            DateTime date1 = new DateTime(start.Ticks - (start.Ticks % TimeSpan.TicksPerSecond), start.Kind);
            
            DataTable dt = new DataTable();
            da = new SqlDataAdapter("Select Top 1 * from RTM_Records left join RTM_SubTask_List on R_SubTask = STL_ID Where R_Start_Date_Time >='"+ date1  +"' and R_User_Name ='" + user + "' and STL_SubTask Like 'Peer Support%'", con);
            da.Fill(dt);
            return dt;
        }

        private void updatePeerTime()
        {
            ds = FetchPeerSupportDataFromLogs();
            TimeSpan taskDiff = TimeSpan.Parse("00:00:00");
            if (ds.Tables["LogPeer"].Rows.Count > 0)
            {
                foreach (DataRow dr in ds.Tables["LogPeer"].Rows)
                {
                    taskDiff = TimeSpan.Parse("00:00:00");
                    DateTime start = Convert.ToDateTime(dr["LA_Start_Date_Time"]);
                    DateTime end = Convert.ToDateTime(dr["LA_CreatedOn"]);
                    taskDiff = TimeSpan.Parse((end - start).ToString(@"hh\:mm\:ss"));
                    string user = dr["LA_User_Name"].ToString();
                    DataTable dt = new DataTable();
                    dt = FetchPeerRecord(start, user);
                    if (dt.Rows.Count > 0)
                    {
                        int recId = Convert.ToInt32(dt.Rows[0]["R_ID"]);

                        cmd = new SqlCommand("UPDATE RTM_Records set R_Duration='"+ taskDiff +"' where R_ID="+ recId +"", con);
                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }
            }
            label1.Text = "Updated";
        }
    }
}
