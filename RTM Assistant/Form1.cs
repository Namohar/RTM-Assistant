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

namespace RTM_Assistant
{
    public partial class Form1 : Form
    {
        SqlConnection con = new SqlConnection(@"Data Source=BLRPRODRTM\RTM_PROD_BLR;Initial Catalog=Real_Time_Metrics;User ID=sa;Password=Prodrtm@123;");
        DataSet ds = new DataSet();
        SqlDataAdapter da;
        SqlCommand cmd;
        int flag = 0;
        int countFlag = 0;
        DateTime recDate;
        DateTime date1;
        DateTime schLogOff;
        string[] ToAddrSuccess = System.Configuration.ConfigurationManager.AppSettings["ToAddressSuccess"].Split(',');
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //UpdateLogoutNew();
            timer1.Enabled = true;
        }

        private void SendMail()
        {
            if (DateTime.Now.DayOfWeek == DayOfWeek.Friday && DateTime.Now.Hour == 10)
            {
                if (flag == 0)
                {
                    try
                    {
                        MailMessage message1 = new MailMessage();
                        SmtpClient smtp = new SmtpClient();

                        message1.From = new MailAddress("BLR-RTM-Server@tangoe.com");
                        //message1.To.Add(new MailAddress("Lokesha.B@tangoe.com"));
                        //foreach (string item in ToAddrSuccess)
                        //{
                            message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                        //}
                        message1.Subject = "RTM Assistant";
                        message1.Body = "This is test mail. Please donot reply";
                        message1.IsBodyHtml = false;

                        smtp.Port = 25;
                        smtp.Host = "outlook-south.tangoe.com";
                        smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                        smtp.EnableSsl = false;
                        //smtp.UseDefaultCredentials = true;
                        //smtp.Credentials = new NetworkCredential("BLR-RTM-Server@tangoe.com", "");
                        smtp.Send(message1);
                        flag = 1;
                    }
                    catch (Exception ex)
                    {
                        WriteToErrorLog(ex.Message, ex.StackTrace, "Error");
                        flag = 0;
                    }
                }
            }

            if (DateTime.Now.DayOfWeek == DayOfWeek.Saturday)
            {
                flag = 0;
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

        private void timer1_Tick(object sender, EventArgs e)
        {
            lblTime.Text = DateTime.Now.ToShortTimeString();
            //SendMail();
            //if (DateTime.Now.DayOfWeek == DayOfWeek.Monday && DateTime.Now.Hour == 3)
            //{
            //    if (countFlag == 0)
            //    {
            //        LastWeekLoginCount();
            //        countFlag = 1;
            //    }
            //    if (DateTime.Now.DayOfWeek == DayOfWeek.Tuesday)
            //    {
            //        flag = 0;
            //    }
            //}
            
        }

        private void notifyIcon1_Click(object sender, EventArgs e)
        {
            this.notifyIcon1.Visible = false;
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
                this.notifyIcon1.Visible = true;
                this.notifyIcon1.ShowBalloonTip(5000);
                this.ShowInTaskbar = false;
            }
        }

        private void tmrLogout_Tick(object sender, EventArgs e)
        {
            try
            {
                //UpdateLogout();
                System.Threading.Thread.Sleep(100);
                UpdateLogoutNew();
            }
            catch (Exception ex)
            {
                WriteToErrorLog(ex.Message, ex.StackTrace, "Error");
            }
        }

        public DataSet GetLastActions()
        {
            da = new SqlDataAdapter("SELECT m1.LA_ID,m1.LA_User_Name,m1.LA_TeamId, m1.LA_Log_Action, m1.LA_Start_Date_Time, m1.LA_Status " +
                            "FROM RTM_Log_Actions m1 LEFT JOIN RTM_Log_Actions m2 "+
                            " ON (m1.LA_User_Name = m2.LA_User_Name AND m1.LA_ID < m2.LA_ID)"+
                            " WHERE m2.LA_ID IS NULL and (m1.LA_Log_Action='Locked' or m1.LA_Log_Action ='Shutdown' or m1.LA_Log_Action='Task Paused')", con);
            da.Fill(ds,"lastAction");
            return ds;
        }

        public DataSet getLastRecord(string user)
        {
            if (ds.Tables.Contains("rec"))
            {
                ds.Tables.Remove(ds.Tables["rec"]);
            }
            da = new SqlDataAdapter("select TOP 1 R_CreatedOn from RTM_Records where R_User_Name ='" + user + "' ORDER BY R_ID DESC", con);
            da.Fill(ds, "rec");
            return ds;
        }

        //Don't Consider
        public void UpdateLogout()
        {
            try
            {
                

                if (ds.Tables.Contains("lastAction"))
                {
                    ds.Tables.Remove(ds.Tables["lastAction"]);
                }

                System.Threading.Thread.Sleep(100);

                ds = GetLastActions();

                if (ds.Tables["lastAction"].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds.Tables["lastAction"].Rows)
                    {
                        int id = Convert.ToInt32(dr["LA_ID"]);
                        int teamId = Convert.ToInt32(dr["LA_TeamId"]);
                        string userName = dr["LA_User_Name"].ToString();
                        string logstatus = dr["LA_Log_Action"].ToString();
                        string lockStatus = dr["LA_Status"].ToString();

                        DateTime date1 = Convert.ToDateTime(dr["LA_Start_Date_Time"]);
                        TimeSpan diff2 = DateTime.Now.Subtract(date1);
                        if (diff2.Hours >= 6 || diff2.Days >= 1)
                        {
                            if (logstatus == "Shutdown" || lockStatus == "Still Locked")
                            {
                                con.Open();
                                cmd = new SqlCommand("UPDATE RTM_Log_Actions SET LA_Log_Action = 'Actual Logout', LA_Status ='' where LA_ID =" + id + "", con);
                                cmd.ExecuteNonQuery();
                                con.Close();

                                ds = GetMultiTasks(userName);

                                if (ds.Tables["tasks"].Rows.Count > 0)
                                {
                                    foreach (DataRow dr2 in ds.Tables["tasks"].Rows)
                                    {
                                        int recordId = Convert.ToInt32(dr2["MT_RecordID"]);
                                        string duration = dr2["MT_TimeSpend"].ToString();
                                        int multitaskId = Convert.ToInt32(dr2["MT_Id"]);

                                        updateTask(recordId, duration, "Completed");

                                        DeleteRecord(multitaskId);

                                    }
                                }
                            }
                            else
                            {
                                con.Open();
                                cmd = new SqlCommand("insert into RTM_Log_Actions (LA_TeamId,LA_User_Name,LA_Log_Action,LA_Start_Date_Time,LA_CreatedOn, LA_Duration) values (" + teamId + ", '" + userName + "', 'Actual Logout', '" + date1.AddMinutes(5) + "', '" + DateTime.Now + "', '')", con);
                                cmd.ExecuteNonQuery();
                                con.Close();

                                ds = GetMultiTasks(userName);

                                if (ds.Tables["tasks"].Rows.Count > 0)
                                {
                                    foreach (DataRow dr2 in ds.Tables["tasks"].Rows)
                                    {
                                        int recordId = Convert.ToInt32(dr2["MT_RecordID"]);
                                        string duration = dr2["MT_TimeSpend"].ToString();
                                        int multitaskId = Convert.ToInt32(dr2["MT_Id"]);

                                        updateTask(recordId, duration, "Completed");

                                        DeleteRecord(multitaskId);

                                    }
                                }
                            }
                        }                       
                       
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToErrorLog(ex.Message, ex.StackTrace, "Error");
            }
    

        }

        private void LastWeekLoginCount()
        {
            if (ds.Tables.Contains("LoginCount"))
            {
                ds.Tables.Remove(ds.Tables["LoginCount"]);
            }
            da = new SqlDataAdapter("select LA_User_Name AS UserName, COUNT(LA_Log_Action) As LoginCount from RTM_Log_actions where LA_Log_Action ='Actual Login' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_Start_Date_Time))) BETWEEN '" + DateTime.Now.AddDays(-(int)DateTime.Now.DayOfWeek - 6).ToShortDateString() + "' and '" + DateTime.Now.AddDays(-(int)DateTime.Now.DayOfWeek).ToShortDateString() + "' GROUP BY LA_User_Name", con);
            da.Fill(ds, "LoginCount");

            Write(ds.Tables["LoginCount"], "" + Directory.GetCurrentDirectory() + "\\LoginCount" + DateTime.Now.AddDays(-(int)DateTime.Now.DayOfWeek - 6).ToString("MM-dd-yyyy") + "To" + DateTime.Now.AddDays(-(int)DateTime.Now.DayOfWeek).ToString("MM-dd-yyyy") + ".txt");
        }

        static void Write(DataTable dt, string outputFilePath)
        {
            int[] maxLengths = new int[dt.Columns.Count];

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                maxLengths[i] = dt.Columns[i].ColumnName.Length;

                foreach (DataRow row in dt.Rows)
                {
                    if (!row.IsNull(i))
                    {
                        int length = row[i].ToString().Length;

                        if (length > maxLengths[i])
                        {
                            maxLengths[i] = length;
                        }
                    }
                }
            }

            using (StreamWriter sw = new StreamWriter(outputFilePath, false))
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    sw.Write(dt.Columns[i].ColumnName.PadRight(maxLengths[i] + 2));
                }

                sw.WriteLine();

                foreach (DataRow row in dt.Rows)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        if (!row.IsNull(i))
                        {
                            sw.Write(row[i].ToString().PadRight(maxLengths[i] + 2));
                        }
                        else
                        {
                            sw.Write(new string(' ', maxLengths[i] + 2));
                        }
                    }

                    sw.WriteLine();
                }

                sw.Close();
            }
        }

        private DataSet GetMultiTasks(string user)
        {
            if (ds.Tables.Contains("tasks"))
            {
                ds.Tables.Remove(ds.Tables["tasks"]);
            }
            da = new SqlDataAdapter("select MT_Id, MT_TimeSpend,MT_RecordID from RTM_Multitasking where MT_UserName = '" + user + "'", con);
            da.Fill(ds, "tasks");
            return ds;
        }

        private void updateTask(int id, string duration, string status)
        {
            con.Open();
            cmd = new SqlCommand("update RTM_Records SET R_Duration ='"+ duration +"', R_Status='"+ status +"' where R_ID ="+ id +"", con);
            cmd.ExecuteNonQuery();
            con.Close();
        }

        private void DeleteRecord(int id)
        {
            con.Open();
            cmd = new SqlCommand("delete from RTM_Multitasking where MT_Id = "+ id +"", con);
            cmd.ExecuteNonQuery();
            con.Close();
        }

        private DataSet GetUsers()
        {
            if (ds.Tables.Contains("users"))
            {
                ds.Tables.Remove(ds.Tables["users"]);
            }
            da = new SqlDataAdapter("select UL_User_Name from RTM_User_List", con);
            da.Fill(ds, "users");
            return ds;
        }

        private DataSet GetLastLog(string user)
        {
            if (ds.Tables.Contains("LastLog"))
            {
                ds.Tables.Remove(ds.Tables["LastLog"]);
            }
            da = new SqlDataAdapter("select TOP 1 * from RTM_Log_Actions, RTM_User_List where LA_User_Name= UL_User_Name and LA_User_Name ='" + user + "' order by LA_ID DESC", con);
            da.Fill(ds, "LastLog");
            return ds;
        }

        private void InsertEarlyLogOff(string empid, string empname, DateTime date, DateTime scheduled, DateTime actual, string totalOfficeHours)
        {
            con.Open();
            cmd = new SqlCommand("Insert into RTM_EarlyLogOffDetails (EL_EmployeeId, EL_User_Name, EL_Date, EL_Scheduled, EL_Actual,EL_CreatedOn, EL_Total_Office_Hours) values ('" + empid + "', '" + empname + "', '" + date + "', '" + scheduled + "', '" + actual + "', '" + DateTime.Now + "', '"+ totalOfficeHours +"') ", con);
            cmd.ExecuteNonQuery();
            con.Close();
        }

        private DataTable  GetLoginTime(string user)
        {
            DataTable dt = new DataTable();
            da = new SqlDataAdapter("SELECT TOP 1 LA_Start_Date_Time From RTM_Log_Actions where LA_User_Name = '" + user + "' and LA_Log_Action = 'Actual Login' Order By LA_Start_Date_Time DESC", con);
            da.Fill(dt);
            return dt;
        }

        private void UpdateLogoutNew()
        {
            ds = GetUsers();

            
            if (ds.Tables["users"].Rows.Count > 0)
            {
                foreach (DataRow dr in ds.Tables["users"].Rows)
                {
                    string userName = dr["UL_User_Name"].ToString();
                    
                    int id=0;
                    int teamId=0;
                    string logstatus=string.Empty ;
                    string lockStatus= string.Empty ;
                    string empId = string.Empty;
                    
                    //Gets Last Log
                    ds = GetLastLog(userName);
                    if (ds.Tables["LastLog"].Rows.Count > 0)
                    {
                        date1 = Convert.ToDateTime(ds.Tables["LastLog"].Rows[0]["LA_Start_Date_Time"]);
                        id = Convert.ToInt32(ds.Tables["LastLog"].Rows[0]["LA_ID"]);
                        teamId = Convert.ToInt32(ds.Tables["LastLog"].Rows[0]["LA_TeamId"]);
                        logstatus = ds.Tables["LastLog"].Rows[0]["LA_Log_Action"].ToString();
                        lockStatus = ds.Tables["LastLog"].Rows[0]["LA_Status"].ToString();
                        empId = ds.Tables["LastLog"].Rows[0]["UL_Employee_Id"].ToString();
                        schLogOff = Convert.ToDateTime(ds.Tables["LastLog"].Rows[0]["UL_SCH_Logout"]);
                    }

                    if (logstatus == "Actual Logout")
                    {
                        continue;
                    }
                    if (logstatus == "Shutdown")
                    {
                        TimeSpan diff = DateTime.Now.Subtract(date1);
                        if (diff.Hours >= 8 || diff.Days >= 1)
                        {
                            con.Open();
                            cmd = new SqlCommand("UPDATE RTM_Log_Actions SET LA_Log_Action = 'Actual Logout', LA_Status ='' where LA_ID =" + id + "", con);
                            cmd.ExecuteNonQuery();
                            con.Close();

                            ds = GetMultiTasks(userName);

                            if (ds.Tables["tasks"].Rows.Count > 0)
                            {
                                foreach (DataRow dr2 in ds.Tables["tasks"].Rows)
                                {
                                    int recordId = Convert.ToInt32(dr2["MT_RecordID"]);
                                    string duration = dr2["MT_TimeSpend"].ToString();
                                    int multitaskId = Convert.ToInt32(dr2["MT_Id"]);

                                    updateTask(recordId, duration, "Completed");

                                    DeleteRecord(multitaskId);

                                }
                            }

                            DataTable dt = new DataTable();
                            dt = GetLoginTime(userName);
                            DateTime logintime = Convert.ToDateTime(dt.Rows[0]["LA_Start_Date_Time"]);
                            //DateTime scheduled = Convert.ToDateTime(logintime.ToShortDateString() + " " + schLogOff.TimeOfDay);
                            TimeSpan datetimediff = TimeSpan.Parse((date1 - logintime).ToString(@"hh\:mm\:ss"));
                            if (datetimediff.Hours < 9 )
                            {
                                InsertEarlyLogOff(empId, userName, logintime, schLogOff, date1, Convert.ToString(datetimediff) );
                            }


                        }
                        continue;
                    }

                    if (lockStatus == "Still Locked")
                    {
                        TimeSpan diff = DateTime.Now.Subtract(date1);
                        if (diff.Hours >= 8 || diff.Days >= 1)
                        {
                            con.Open();
                            cmd = new SqlCommand("UPDATE RTM_Log_Actions SET LA_Log_Action = 'Actual Logout', LA_Status ='' where LA_ID =" + id + "", con);
                            cmd.ExecuteNonQuery();
                            con.Close();

                            DataTable dt = new DataTable();
                            dt = GetLoginTime(userName);
                            DateTime logintime = Convert.ToDateTime(dt.Rows[0]["LA_Start_Date_Time"]);
                            //DateTime scheduled = Convert.ToDateTime(logintime.ToShortDateString() + " " + schLogOff.TimeOfDay);
                            TimeSpan datetimediff = TimeSpan.Parse((date1 - logintime).ToString(@"hh\:mm\:ss"));
                            if (datetimediff.Hours < 9)
                            {
                                InsertEarlyLogOff(empId, userName, logintime, schLogOff, date1, Convert.ToString(datetimediff));
                            }

                            ds = GetMultiTasks(userName);

                            if (ds.Tables["tasks"].Rows.Count > 0)
                            {
                                foreach (DataRow dr2 in ds.Tables["tasks"].Rows)
                                {
                                    int recordId = Convert.ToInt32(dr2["MT_RecordID"]);
                                    string duration = dr2["MT_TimeSpend"].ToString();
                                    int multitaskId = Convert.ToInt32(dr2["MT_Id"]);

                                    updateTask(recordId, duration, "Completed");

                                    DeleteRecord(multitaskId);

                                }
                            }
                        }
                 
                        continue;
                    }
                    
                    //Gets Last Record
                    ds= getLastRecord(userName);
                    if (ds.Tables["rec"].Rows.Count > 0)
                    {
                        recDate = Convert.ToDateTime(ds.Tables["rec"].Rows[0]["R_CreatedOn"]);
                    }
                   
                    if (date1 > recDate)
                    {
                        TimeSpan diff = DateTime.Now.Subtract(date1);
                        if (diff.Hours >= 8 || diff.Days >= 1)
                        {
                            con.Open();
                            if (logstatus == "Shutdown")
                            {
                                cmd = new SqlCommand("UPDATE RTM_Log_Actions SET LA_Log_Action = 'Actual Logout', LA_Status ='' where LA_ID =" + id + "", con);
                                cmd.ExecuteNonQuery();
                                con.Close();

                                DataTable dt = new DataTable();
                                dt = GetLoginTime(userName);
                                DateTime logintime = Convert.ToDateTime(dt.Rows[0]["LA_Start_Date_Time"]);
                                //DateTime scheduled = Convert.ToDateTime(logintime.ToShortDateString() + " " + schLogOff.TimeOfDay);
                                TimeSpan datetimediff = TimeSpan.Parse((date1 - logintime).ToString(@"hh\:mm\:ss"));
                                if (datetimediff.Hours < 9)
                                {
                                    InsertEarlyLogOff(empId, userName, logintime, schLogOff, date1, Convert.ToString(datetimediff));
                                }
                            }
                            else if (lockStatus == "Still Locked")
                            {
                                cmd = new SqlCommand("UPDATE RTM_Log_Actions SET LA_Log_Action = 'Actual Logout', LA_Status ='' where LA_ID =" + id + "", con);
                                cmd.ExecuteNonQuery();
                                con.Close();

                                DataTable dt = new DataTable();
                                dt = GetLoginTime(userName);
                                DateTime logintime = Convert.ToDateTime(dt.Rows[0]["LA_Start_Date_Time"]);
                                //DateTime scheduled = Convert.ToDateTime(logintime.ToShortDateString() + " " + schLogOff.TimeOfDay);
                                TimeSpan datetimediff = TimeSpan.Parse((date1 - logintime).ToString(@"hh\:mm\:ss"));
                                if (datetimediff.Hours < 9)
                                {
                                    InsertEarlyLogOff(empId, userName, logintime, schLogOff, date1, Convert.ToString(datetimediff));
                                }
                            }
                            else
                            {
                                cmd = new SqlCommand("insert into RTM_Log_Actions (LA_TeamId,LA_User_Name,LA_Log_Action,LA_Start_Date_Time,LA_CreatedOn, LA_Duration) values (" + teamId + ", '" + userName + "', 'Actual Logout', '" + date1.AddMinutes(5) + "', '" + date1.AddMinutes(5) + "', '')", con);
                                cmd.ExecuteNonQuery();
                                con.Close();

                                DataTable dt = new DataTable();
                                dt = GetLoginTime(userName);
                                DateTime logintime = Convert.ToDateTime(dt.Rows[0]["LA_Start_Date_Time"]);
                                //DateTime scheduled = Convert.ToDateTime(logintime.ToShortDateString() + " " + schLogOff.TimeOfDay);
                                TimeSpan datetimediff = TimeSpan.Parse((date1 - logintime).ToString(@"hh\:mm\:ss"));
                                if (datetimediff.Hours < 9)
                                {
                                    InsertEarlyLogOff(empId, userName, logintime, schLogOff, date1.AddMinutes(5), Convert.ToString(datetimediff));
                                }
                            }

                            ds = GetMultiTasks(userName);

                            if (ds.Tables["tasks"].Rows.Count > 0)
                            {
                                foreach (DataRow dr2 in ds.Tables["tasks"].Rows)
                                {
                                    int recordId = Convert.ToInt32(dr2["MT_RecordID"]);
                                    string duration = dr2["MT_TimeSpend"].ToString();
                                    int multitaskId = Convert.ToInt32(dr2["MT_Id"]);

                                    updateTask(recordId, duration, "Completed");

                                    DeleteRecord(multitaskId);

                                }
                            }
                        }
                    }
                    else if(recDate > date1)
                    {
                        TimeSpan diff = DateTime.Now.Subtract(recDate);
                        if (diff.Hours >= 8 || diff.Days >= 1)
                        {
                            con.Open();
                            cmd = new SqlCommand("insert into RTM_Log_Actions (LA_TeamId,LA_User_Name,LA_Log_Action,LA_Start_Date_Time,LA_CreatedOn, LA_Duration) values (" + teamId + ", '" + userName + "', 'Actual Logout', '" + recDate.AddMinutes(5) + "', '" + recDate.AddMinutes(5) + "', '')", con);
                            cmd.ExecuteNonQuery();
                            con.Close();

                            DataTable dt = new DataTable();
                            dt = GetLoginTime(userName);
                            DateTime logintime = Convert.ToDateTime(dt.Rows[0]["LA_Start_Date_Time"]);
                            //DateTime scheduled = Convert.ToDateTime(logintime.ToShortDateString() + " " + schLogOff.TimeOfDay);
                            TimeSpan datetimediff = TimeSpan.Parse((date1 - logintime).ToString(@"hh\:mm\:ss"));
                            if (datetimediff.Hours < 9)
                            {
                                InsertEarlyLogOff(empId, userName, logintime, schLogOff, recDate.AddMinutes(5), Convert.ToString(datetimediff));
                            }

                            ds = GetMultiTasks(userName);

                            if (ds.Tables["tasks"].Rows.Count > 0)
                            {
                                foreach (DataRow dr2 in ds.Tables["tasks"].Rows)
                                {
                                    int recordId = Convert.ToInt32(dr2["MT_RecordID"]);
                                    string duration = dr2["MT_TimeSpend"].ToString();
                                    int multitaskId = Convert.ToInt32(dr2["MT_Id"]);

                                    updateTask(recordId, duration, "Completed");

                                    DeleteRecord(multitaskId);

                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
