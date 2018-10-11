using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;

namespace RTMTimeDate
{
    public partial class Form1 : Form
    {
        int flag = 0;
        SqlConnection con = new SqlConnection(@"Data Source=10.55.5.40,1433;Initial Catalog=RTM_Global_Test;User ID=PRODRTMDB;Password=Prodrtm@123;");

        //Dev
        //SqlConnection con = new SqlConnection(@"Data Source=10.55.5.40,1433;Initial Catalog=Real_Time_Metrics_Dev;User ID=PRODRTMDB;Password=Prodrtm@123;");
        SqlCommand cmd;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Enabled = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            System.Threading.Thread.Sleep(1000);
            
            TimeDate();
        }

        private void notifyIconTimeDate_Click(object sender, EventArgs e)
        {
            this.notifyIconTimeDate.Visible = false;
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.notifyIconTimeDate.Visible = true;
                this.notifyIconTimeDate.ShowBalloonTip(5000);
                this.ShowInTaskbar = false;
            }
        }

        private void UpdateAllTimeDate()
        {
            
            using (cmd = new SqlCommand("update RTM_Records set R_TimeDate= R_Start_Date_Time where R_TimeDate is null", con))
            {                
                cmd.CommandTimeout = 0;
                if (con.State == ConnectionState.Open)
                {
                    con.Close();                    
                }
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }
            

            
            using (cmd = new SqlCommand("update RTM_Log_Actions set LA_TimeDate= LA_Start_Date_Time where LA_TimeDate is null", con))
            {
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }
                cmd.CommandTimeout = 0;
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }

            using (cmd = new SqlCommand("update RTM_IPVDetails set TimeDate= StartTime where TimeDate is null", con))
            {
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }
                cmd.CommandTimeout = 0;
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }
            
        }

        private void UpdateNightShiftTimeDate()
        {

            using (cmd = new SqlCommand("update rec set rec.R_TimeDate = DATEADD(day,-1,rec.R_Start_Date_Time) from RTM_Records rec, RTM_Team_List tl, RTM_User_List where rec.R_TeamId = tl.T_ID and R_User_Name = UL_User_Name and CONVERT(char(10), rec.R_Start_Date_Time, 108) BETWEEN '00:00:00' and '07:00:00'  and tl.T_Location='IND' and R_System is null and R_Task <>0 and R_TimeDate = R_Start_Date_Time and CONVERT(char(10), UL_SCH_Logout, 108) between '00:00:00' and '07:00:00'", con))
            {
                cmd.CommandTimeout = 0;
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }

            using (cmd = new SqlCommand("update rec set rec.LA_TimeDate = DATEADD(day,-1,rec.LA_Start_Date_Time) from RTM_Log_Actions rec, RTM_Team_List tl, RTM_User_List where rec.LA_TeamId = tl.T_ID and LA_User_Name = UL_User_Name and CONVERT(char(10), rec.LA_Start_Date_Time, 108) BETWEEN '00:00:00' and '07:00:00'  and tl.T_Location='IND' and  LA_TimeDate = LA_Start_Date_Time and CONVERT(char(10), UL_SCH_Logout, 108) between '00:00:00' and '07:00:00'", con))
            {
                cmd.CommandTimeout = 0;
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }

            using (cmd = new SqlCommand("update rec set rec.TimeDate = DATEADD(day,-1,rec.StartTime) from RTM_IPVDetails rec, RTM_Team_List tl, RTM_User_List where rec.Team_Id = tl.T_ID and UserName = UL_User_Name and CONVERT(char(10), rec.StartTime, 108) BETWEEN '00:00:00' and '07:00:00'  and tl.T_Location='IND' and TimeDate = StartTime and CONVERT(char(10), UL_SCH_Logout, 108) between '00:00:00' and '07:00:00'", con))
            {
                cmd.CommandTimeout = 0;
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }
        }

        private void TimeDate()
        {
            try
            {
                UpdateAllTimeDate();

                if (DateTime.Now.Hour == 6)
                {
                    if (flag == 0)
                    {
                        flag = 1;
                        UpdateNightShiftTimeDate();
                    }
                }

                if (DateTime.Now.Hour == 7)
                {
                    flag = 0;
                }

            }
            catch (Exception ex)
            {
                WriteToErrorLog(ex.Message, ex.StackTrace, "Error");
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }
            }

        }

        private void WriteToErrorLog(string msg, string stkTrace, string title)
        {
            
            if (!System.IO.Directory.Exists(Application.StartupPath + "\\Errors\\"))
            {
                System.IO.Directory.CreateDirectory(Application.StartupPath + "\\Errors\\");
            }

            FileStream fs = new FileStream(Application.StartupPath + "\\Errors\\errlog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamWriter s = new StreamWriter(fs);
            s.Close();
            fs.Close();

            FileStream fs1 = new FileStream(Application.StartupPath + "\\Errors\\errlog.txt", FileMode.Append, FileAccess.Write);
            StreamWriter s1 = new StreamWriter(fs1);
            s1.Write("Title: " + title + Environment.NewLine);
            s1.Write("Message: " + msg + Environment.NewLine);
            s1.Write("StackTrace: " + stkTrace + Environment.NewLine);
            s1.Write("Date/Time: " + DateTime.Now.ToString() + Environment.NewLine);
            s1.Write("================================================" + Environment.NewLine);
            s1.Close();
            fs1.Close();
        }

        private void CheckClockin()
        {
            
        }
    }
}
