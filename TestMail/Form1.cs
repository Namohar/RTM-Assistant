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

namespace TestMail
{
    public partial class txtPort : Form
    {
        SqlConnection con = new SqlConnection(@"Data Source=BLRPRODRTM\RTM_PROD_BLR;Initial Catalog=Real_Time_Metrics;User ID=sa;Password=Prodrtm@123;");

        SqlConnection globalCon = new SqlConnection(@"Data Source=10.55.5.40,1433;Initial Catalog=Real_Time_Metrics_Dev;User ID=PRODRTMDB;Password=Prodrtm@123;");
        DataSet ds = new DataSet();
        SqlDataAdapter da;
        SqlCommand cmd;
        StringBuilder myBuilder = new StringBuilder();
        public txtPort()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            sendMail();
            //SendMail();
            //ds = GetEarlyLogOffDetails();

            //if (ds.Tables["early"].Rows.Count > 0)
            //{
            //    foreach (DataRow dr in ds.Tables["early"].Rows)
            //    {
            //        int id = Convert.ToInt32(dr["EL_ID"]);
            //        string user = dr["EL_User_Name"].ToString();
            //        DateTime dtLogoff = Convert.ToDateTime(dr["EL_Actual"]);

            //        DataTable dt = new DataTable();
            //        dt = GetLoginTime(user, dtLogoff);
            //        if (dt.Rows.Count > 0)
            //        {
            //            DateTime logintime = Convert.ToDateTime(dt.Rows[0]["LA_Start_Date_Time"]);
            //            var datetimediff = (dtLogoff - logintime).ToString(@"hh\:mm\:ss");

            //            UpdateEarlyLogoff(id, datetimediff);
            //        }
                    
            //    }
            //}
        }

        private DataSet GetEarlyLogOffDetails()
        {
            da = new SqlDataAdapter("select EL_ID, EL_User_Name, EL_Actual from RTM_EarlyLogOffDetails", con);
            da.Fill(ds, "early");
            return ds;
        }

        private DataTable GetLoginTime(string user, DateTime logoff)
        {
            DataTable dt = new DataTable();
            da = new SqlDataAdapter("SELECT TOP 1 LA_Start_Date_Time From RTM_Log_Actions where LA_User_Name = '" + user + "' and LA_Log_Action = 'Actual Login' and LA_Start_Date_Time<='"+ logoff +"' Order By LA_Start_Date_Time DESC", con);
            da.Fill(dt);
            return dt;
        }

        private void UpdateEarlyLogoff(int id, string duration)
        {
            cmd = new SqlCommand("UPDATE RTM_EarlyLogOffDetails SET EL_Total_Office_Hours='" + duration + "' where EL_ID='"+ id +"'", con);
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }

        private DataSet GetTotalEstimateExpected(int TID)
        {
            if (ds.Tables.Contains("expected"))
            {
                ds.Tables.Remove(ds.Tables["expected"]);
            }
            da = new SqlDataAdapter("select COUNT(Distinct EST_UserName) as [User Count], "+
                                "CONVERT(varchar(10), sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/3600) +':'+ "+
                                "CONVERT(varchar(10),(sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/60)%60) +':'+ "+
                                "CONVERT(varchar(10),(sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/60)%60)))%60) as [Estimated Total Duration], "+
                                "CONVERT(varchar(10),COUNT(Distinct EST_UserName) * 8) +':00:00' as [Expected Total Duration] "+
                                "from RTM_Estimation where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, EST_Date)))='8/28/2015' and EST_TeamId='13'", con);
            da.Fill(ds, "expected");
            return ds;
        }

        private DataSet GetActualEstimate(int TID)
        {
            if (ds.Tables.Contains("actual"))
            {
                ds.Tables.Remove(ds.Tables["actual"]);
            }
            da = new SqlDataAdapter("select EST_UserName as [User Name], "+
                                    "CONVERT(varchar(10), sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/3600) +':'+ "+
                                    "CONVERT(varchar(10),(sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/60)%60) +':'+ "+
                                    "CONVERT(varchar(10),(sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/60)%60)))%60) as [Estimated Total Duration] "+
                                    "from RTM_Estimation where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, EST_Date)))='8/28/2015' and EST_TeamId='13' Group By EST_UserName HAVING sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/3600 < 8", con);
            da.Fill(ds, "actual");
            return ds;
        }

        private void SendMail()
        {
            try
            {
                MailMessage message1 = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                //ds = GetTotalEstimateExpected(10);
                //ds = GetActualEstimate(10);
                message1.From = new MailAddress(txtFrom.Text);
                message1.To.Add(new MailAddress(txtTo.Text));
                //message1.To.Add(new MailAddress("Subhankar.Brahma@tangoe.com"));
                message1.Subject = "Test Mail";
                //getHTML(ds);
                //StringBuilder sb = new StringBuilder();
                //sb.AppendLine("Hi All,");
                //sb.AppendLine("");
                //sb.AppendLine("Please find the attached report for today's Resource Utilization Estimate.");
                //sb.AppendLine("");
                ////sb.AppendLine(myBuilder.ToString());   //here I want the data to       display in table format
                //sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                //sb.AppendLine("");

                message1.Body = "Teasting";// sb.ToString();
                message1.IsBodyHtml = true;
                smtp.Port = Convert.ToInt32(txtPort1.Text);
                smtp.Host = txtHost.Text;
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);

                label1.Text = "Mail Sent";
            }
            catch (Exception ex)
            {

                label1.Text = ex.Message;
                label2.Text = ex.InnerException + Environment.NewLine + "   Stack Trace" + ex.StackTrace;
            }
           

        }

        private string getHTML(DataSet ds1)
        {
            

            myBuilder.Append("<table border='1' cellpadding='5' cellspacing='0' ");
            myBuilder.Append("style='border: solid 1px Silver; font-size: x-small;'>");

            myBuilder.Append("<tr align='left' valign='top'>");
            foreach (DataColumn myColumn in ds1.Tables["expected"].Columns)
            {
                myBuilder.Append("<td align='left' valign='top' bgcolor='#FFFF00'>");
                myBuilder.Append("<B />" + myColumn.ColumnName);
                myBuilder.Append("</td>");
            }
            myBuilder.Append("</tr>");

            foreach (DataRow myRow in ds1.Tables["expected"].Rows)
            {
                myBuilder.Append("<tr align='left' valign='top'>");
                foreach (DataColumn myColumn in ds1.Tables["expected"].Columns)
                {
                    myBuilder.Append("<td align='left' valign='top'>");
                    myBuilder.Append(myRow[myColumn.ColumnName].ToString());
                    myBuilder.Append("</td>");
                }
                myBuilder.Append("</tr>");
            }
            myBuilder.Append("<tr align='left' valign='top'>");
            foreach (DataColumn myColumn in ds1.Tables["actual"].Columns)
            {
                myBuilder.Append("<td align='left' valign='top' bgcolor='#FFFF00'>");
                myBuilder.Append("<B />" + myColumn.ColumnName);
                myBuilder.Append("</td>");
            }
            myBuilder.Append("</tr>");

            foreach (DataRow myRow in ds1.Tables["actual"].Rows)
            {
                myBuilder.Append("<tr align='left' valign='top'>");
                foreach (DataColumn myColumn in ds1.Tables["actual"].Columns)
                {
                    myBuilder.Append("<td align='left' valign='top'>");
                    myBuilder.Append(myRow[myColumn.ColumnName].ToString());
                    myBuilder.Append("</td>");
                }
                myBuilder.Append("</tr>");
            }
            myBuilder.Append("</table>");

            return myBuilder.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           SendMail();
            //SendLateLoginDetails(DateTime.Now.ToShortDateString(), 6, "07:00", "08:30");//WEM
           // SendLateLoginDetails(DateTime.Now.ToShortDateString(), 7, "07:00", "08:30");//Invoice3
            //SendLateLoginDetails(DateTime.Now.ToShortDateString(), 8, "07:00", "08:30");//Invoice2
           // SendLateLoginDetails(DateTime.Now.ToShortDateString(), 11, "07:00", "08:30");//Audit
           // SendLateLoginDetails(DateTime.Now.ToShortDateString(), 13, "07:00", "08:30");//Implementation
            //SendLateLoginDetails(DateTime.Now.ToShortDateString(), 14, "07:00", "08:30");//Catelog
            //SendLateLoginDetails(DateTime.Now.ToShortDateString(), 18, "07:00", "08:30");//Ops Support
           // SendLateLoginDetails(DateTime.Now.ToShortDateString(), 9, "08:00", "10:00"); // QC
          //  SendLateLoginDetails(DateTime.Now.ToShortDateString(), 10, "09:56", "14:55");//Onboarding



            //TimeSpan span = (Convert.ToDateTime(DateTime.Now.AddMilliseconds(100)) - Convert.ToDateTime(DateTime.Now));
            //int ms = (int)span.TotalMilliseconds;
        }


















        private void SendLateLoginDetails(string date, int teamid, string fromTime, string toTime)
        {
           DataTable dt = new DataTable();
           if (teamid == 8)
           {
               da = new SqlDataAdapter("select D_UserName as [User Name],UL_Employee_Id as [Employee Id], CONVERT(VARCHAR(10), D_Date, 101) as [Date],CONVERT(VARCHAR(10), D_SLogin, 108) as [Sceduled Login], CONVERT(VARCHAR(10), D_Date, 108) as [Actual Login], D_Duration as [Delayed Login], D_Reason as [Delay Reason] " +
                   "from dbo.RTM_DelayedLogInOff, dbo.RTM_User_List where D_UserName = UL_User_Name and convert(char(5), dateadd(minute, 60 + (datediff(minute, 0, D_Date) / 60) * 60, 0), 108) between '" + fromTime + "' and '" + toTime + "' " +
                    "and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, D_Date))) = '" + date + "' and (D_Team_Id = 8 or D_Team_Id = 14) and D_Type = 'In' Order By D_Reason, D_UserName", con);
           }
           else
           {
               da = new SqlDataAdapter("select D_UserName as [User Name],UL_Employee_Id as [Employee Id], CONVERT(VARCHAR(10), D_Date, 101) as [Date],CONVERT(VARCHAR(10), D_SLogin, 108) as [Sceduled Login], CONVERT(VARCHAR(10), D_Date, 108) as [Actual Login], D_Duration as [Delayed Login], D_Reason as [Delay Reason] " +
                   "from dbo.RTM_DelayedLogInOff, dbo.RTM_User_List where D_UserName = UL_User_Name and convert(char(5), dateadd(minute, 60 + (datediff(minute, 0, D_Date) / 60) * 60, 0), 108) between '" + fromTime + "' and '" + toTime + "' " +
                    "and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, D_Date))) = '" + date + "' and D_Team_Id = " + teamid + " and D_Type = 'In' Order By D_Reason, D_UserName", con);
           }
            
            da.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                getDelayHTML(dt);

                StringBuilder sb = new StringBuilder();

                sb.AppendLine("");
                sb.AppendLine(myBuilder.ToString());
                sb.AppendLine("");//here I want the data to       display in table format
                sb.AppendLine("");
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                sb.AppendLine("");

                MailMessage message1 = new MailMessage();
                SmtpClient smtp = new SmtpClient();

                //message1.From = new MailAddress("BLR-RTM-Server@tangoe.com");
                //message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                ////message1.To.Add(new MailAddress("Sriram.Krishnan@tangoe.com"));
                //message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                //message1.To.Add(new MailAddress("Lokesha.B@tangoe.com"));
                if (teamid == 1) // OrderDesk
                {
                    message1.From = new MailAddress("BLR-RTM-Server@tangoe.com");
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Sriram.Krishnan@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Fulfillment Team -  Late Login Summary (" + DateTime.Now.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 6) //Wem
                {
                    message1.From = new MailAddress("BLR-RTM-Server@tangoe.com");
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Arjun.Nagraj@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "WEM Team -  Late Login Summary (" + DateTime.Now.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 7) // Invoice 3
                {
                    message1.From = new MailAddress("BLR-RTM-Server@tangoe.com");
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Balaji.Nagabhushanam@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Invoice 3 Team -  Late Login Summary (" + DateTime.Now.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 8) //Invoice 2
                {
                    message1.From = new MailAddress("BLR-RTM-Server@tangoe.com");
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Vikas.Vyas@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Invoice 2 Team -  Late Login Summary (" + DateTime.Now.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 9) // Quality Check
                {
                    message1.From = new MailAddress("BLR-RTM-Server@tangoe.com");
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Quality Check Team -  Late Login Summary (" + DateTime.Now.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 10) //Onboarding
                {
                    message1.From = new MailAddress("BLR-RTM-Server@tangoe.com");
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Meena.Lakshmi@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Onboarding Team -  Late Login Summary (" + DateTime.Now.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 11) //Audit & Optimize
                {
                    message1.From = new MailAddress("BLR-RTM-Server@tangoe.com");
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Vinith.Bekal@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Audit & Optimize Team -  Late Login Summary (" + DateTime.Now.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 13) // Implementation
                {
                    message1.From = new MailAddress("BLR-RTM-Server@tangoe.com");
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Meena.Lakshmi@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Implementation Team -  Late Login Summary (" + DateTime.Now.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 14) //Catalog & Mapping
                {
                    message1.From = new MailAddress("BLR-RTM-Server@tangoe.com");
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Vikas.Vyas@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Catalog & Mapping Team -  Late Login Summary (" + DateTime.Now.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 18) //Ops Support
                {
                    message1.From = new MailAddress("BLR-RTM-Server@tangoe.com");
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Meena.Lakshmi@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Ops Support Team -  Late Login Summary (" + DateTime.Now.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }


                message1.Body = sb.ToString();
                message1.IsBodyHtml = true;
                smtp.Port = 25;
                smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }

        }

        private string getDelayHTML(DataTable dt)
        {
            myBuilder = new StringBuilder();

            myBuilder.Append("<table border='1' cellpadding='5' cellspacing='0' ");
            myBuilder.Append("style='border: solid 1px Silver; font-size: x-small;'>");

            myBuilder.Append("<tr align='left' valign='top'>");
            foreach (DataColumn myColumn in dt.Columns)
            {
                myBuilder.Append("<td align='left' valign='top' bgcolor='#FFFF00'>");
                myBuilder.Append("<B />" + myColumn.ColumnName);
                myBuilder.Append("</td>");
            }
            myBuilder.Append("</tr>");

            foreach (DataRow myRow in dt.Rows)
            {
                myBuilder.Append("<tr align='left' valign='top'>");
                foreach (DataColumn myColumn in dt.Columns)
                {
                    myBuilder.Append("<td align='left' valign='top'>");
                    myBuilder.Append(myRow[myColumn.ColumnName].ToString());
                    myBuilder.Append("</td>");
                }
                myBuilder.Append("</tr>");
            }
            myBuilder.Append("</table>");

            return myBuilder.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var sourceLoc = @"\\files\Shares\HRIS Data\Production";
            DataTable data = new DataTable();
            DataTable dtEmp = new DataTable();
            var recent = new DirectoryInfo(sourceLoc).GetDirectories()
                       .OrderByDescending(d => d.LastWriteTimeUtc).First();

            var recentfolder = recent.FullName;

            var directory = new DirectoryInfo(recentfolder);
            //var myFile = directory.GetFiles()
            // .OrderByDescending(f => f.LastWriteTime)
            // .First();

            var myFile = directory.GetFiles();

            foreach (var file in myFile)
            {
                if (file.FullName.Contains("1TRQ_IMPORT_EMP_MGR"))
                {
                    var fileName = file.FullName;

                    var filename1 = fileName;
                    var reader = ReadAsLines(filename1);

                    

                    //this assume the first record is filled with the column names
                    var headers = reader.First().Split('|');
                    foreach (var header in headers)
                        data.Columns.Add(header);

                    var records = reader.Skip(1);
                    foreach (var record in records)
                        data.Rows.Add(record.Split('|'));

                    if (data.Rows.Count > 0)
                    {
                        DeactivateAllEmployees();
                        dtEmp = CheckExistingEmployee();
                        foreach (DataRow drRow in data.Rows)
                        {

                           // DataRow[] dr = dtEmp.Select("MUL_EmployeeId = '" + drRow["employee_id"].ToString() + "'");
                            var result = dtEmp.AsEnumerable().Where(drCheck => drCheck.Field<string>("MUL_EmployeeId") == drRow["employee_id"].ToString()).ToList();
                            if (result.Count > 0)
                            {
                                UpdateDepartment(drRow["employee_id"].ToString(), drRow["department_id"].ToString());
                            }
                            else
                            {
                                if (drRow["department_id"].ToString() == "Unknown")
                                {
                                    continue;
                                }

                                SqlParameter[] parameters = new SqlParameter[]
                                {
                                    new SqlParameter("@empId", drRow["employee_id"].ToString()),
                                    new SqlParameter("@first", drRow["first_name"].ToString()),
                                    new SqlParameter("@last", drRow["last_name"].ToString()),
                                    new SqlParameter("@emailId", drRow["email_address"].ToString()),
                                    new SqlParameter("@managerId", drRow["department_id"].ToString()),
                                    new SqlParameter("@createdOn", DateTime.Now),
                                    new SqlParameter("@status", 1)
                                };
                                string sQuery = "insert into RTM_Master_UserList (MUL_EmployeeId,MUL_FirstName,MUL_LastName,MUL_EmailId,MUL_ManagerID,MUL_CreatedOn,MUL_ActiveStatus) "+
                                                 "values(@empId, @first,@last,@emailId,@managerId,@createdOn,@status)";
                                using (cmd = new SqlCommand())
                                {
                                    cmd.Parameters.AddRange(parameters);
                                    cmd.CommandText = sQuery;
                                    cmd.CommandType = CommandType.Text;
                                    cmd.Connection = globalCon;
                                    globalCon.Open();
                                    cmd.ExecuteNonQuery();
                                    globalCon.Close();
                                }
                            }
                        }
                    }

                    UpdateManagerEmail();

                    return;
                }
            }
        }

        static IEnumerable<string> ReadAsLines(string filename)
        {
            using (var reader = new StreamReader(filename))
                while (!reader.EndOfStream)
                    yield return reader.ReadLine();
        }

        public DataTable CheckExistingEmployee()
        {
            DataTable dt = new DataTable();

            using (da = new SqlDataAdapter("select * from RTM_Master_UserList", globalCon))
            {
                da.Fill(dt);
            }
            return dt;
        }

        private void DeactivateAllEmployees()
        {
            using (cmd = new SqlCommand("update RTM_Master_UserList set MUL_ActiveStatus =0", globalCon))
            {
                globalCon.Open();
                cmd.ExecuteNonQuery();
                globalCon.Close();
            }
        }

        private void UpdateDepartment(string employeeId, string managerId)
        {
            using (cmd = new SqlCommand("update RTM_Master_UserList set MUL_ActiveStatus =1, MUL_ManagerID='" + managerId + "' where MUL_EmployeeId='" + employeeId + "'", globalCon))
            {
                globalCon.Open();
                cmd.ExecuteNonQuery();
                globalCon.Close();
            }
        }

        private void UpdateManagerEmail()
        {
            string sQuery = "Update ul Set ul.MUL_ManagerEmail_id = mu.MUL_EmailId From RTM_Master_UserList ul Inner join RTM_Master_UserList mu on ul.MUL_ManagerID = mu.MUL_EmployeeId";

            using (cmd = new SqlCommand(sQuery, globalCon))
            {
                globalCon.Open();
                cmd.ExecuteNonQuery();
                globalCon.Close();
            }
        }

        private void sendMail()
        {
            MailMessage message1 = new MailMessage();
            SmtpClient smtp = new SmtpClient();
            message1.From = new MailAddress("Lokesha.B@tangoe.com");
            message1.Body = "Testing";
            message1.Subject = "Testmail";
            message1.To.Add(new MailAddress("Lokesha.B@tangoe.com"));
            message1.Headers.Set("Date", "09 Jan 1999 17:23:42 -0400");
            SmtpClient smtpClient = new SmtpClient("mail.north.tangoe.com");
            smtpClient.UseDefaultCredentials = false;
            NetworkCredential credentials = new NetworkCredential("Lokesha.B", "Lokeshmca11");
            
            smtpClient.Credentials = credentials;
            smtpClient.Send(message1);
        }
    }
}
