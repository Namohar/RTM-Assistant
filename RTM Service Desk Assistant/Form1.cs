using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;

namespace RTM_Service_Desk_Assistant
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataTable dtUsers = new DataTable();
        DataTable dtRTMClients = new DataTable();
        DataTable dtAgentState = new DataTable();
        
        string conString = ConfigurationManager.AppSettings["conString"].ToString();
        private static TimeZoneInfo INDIAN_ZONE = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");

        private void Form1_Load(object sender, EventArgs e)
        {
           // GetVCCAPI_Data();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetVCCAPI_Data();
        }
        private void GetVCCAPI_Data()
        {
            try
            {
                DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
                label1.Text = "Process started Please wait....";
                label2.Text = "Start Time - " + DateTime.Now;
                label3.Text = "";
                com.incontact.login.inSideWS objRep = new com.incontact.login.inSideWS();

                DateTime start = Convert.ToDateTime(indianTime.AddDays(-(int)DateTime.Now.DayOfWeek - 7).ToShortDateString()+" 0:00:00 AM"); // Convert.ToDateTime(DateTime.Now.AddDays(-(int)DateTime.Now.DayOfWeek - 7)).ToUniversalTime();
                DateTime end = Convert.ToDateTime(indianTime.AddDays(-(int)DateTime.Now.DayOfWeek - 1).ToShortDateString() +" 11:59:59 PM"); // Convert.ToDateTime(DateTime.Now.AddDays(-(int)DateTime.Now.DayOfWeek - 1)).ToUniversalTime();


                //DateTime start = Convert.ToDateTime(DateTime.Now.AddDays(-1)).ToUniversalTime();
                //DateTime end = Convert.ToDateTime(DateTime.Now).ToUniversalTime();
                //start = Convert.ToDateTime("07/28/2017");
                //end = Convert.ToDateTime("07/29/2017");
                var dates = Enumerable.Range(0, 1 + end.Subtract(start).Days)
                              .Select(offset => start.AddDays(offset))
                              .ToArray();
               
                com.incontact.login.inCredentials objinCred = new com.incontact.login.inCredentials();
                objinCred.busNo = 4593343;
                objinCred.password = "59D4CE34-672B-4915-AA66-FC9D742461AB";

                objinCred.partnerPassword = "";
                objinCred.timeZoneName = "";
                objinCred.DialerDSNName = "";
                objinCred.DialerDSNUserID = "";
                objinCred.DialerDSNPassword = "";

                objRep.inCredentialsValue = objinCred;

                //var myXML = objRep.DataDownloadReport_Run(16, start, end);

                //Expanded Call details
                var myXML = objRep.DataDownloadReport_Run(16, start, end);
                DataTable dt = new DataTable();
                dt = myXML.Tables[0];

                //Agent State log
                var agentStateXML = objRep.DataDownloadReport_Run(350083, start, end);

                dtAgentState = agentStateXML.Tables[0];

                dtUsers = GetUsers();
                dtRTMClients = GetClients();
                if (dtUsers.Rows.Count > 0)
                {
                    DataTable dtResult = new DataTable();
                    if (dt.Rows.Count > 0)
                    {
                        foreach (var selDate in dates)
                        {
                            string date = selDate.ToString("MM/dd/yyyy");
                            foreach (DataRow drUser in dtUsers.Rows)
                            {
                                string empId = drUser["UL_Employee_Id"].ToString().Trim();
                                string user = drUser["UL_User_Name"].ToString().Trim();
                                string userName = drUser["UL_System_User_Name"].ToString();
                                userName = userName.Substring(5, userName.Length - 5);
                                userName = userName.Replace('.', ' ');
                                if (userName == "ashwini1")
                                {
                                    userName = "Ashwini Raghunathan";
                                }
                                else if (userName == "c nikitha")
                                {
                                    userName = "Nikitha C";
                                }
                                else if (userName == "Dinesh S")
                                {
                                    userName = "Dinesh Babu S";
                                }
                                else if (userName == "Edward A")
                                {
                                    userName = "Edward Naveen A";
                                }
                                else if (userName == "faizal a") //
                                {
                                    userName = "Faizal Rahman";
                                }
                                else if (userName == "fantin j") //
                                {
                                    userName = "Fantin Cyril";
                                }
                                else if (userName == "sandeep v") //
                                {
                                    userName = "Sandeep Rathnakaran";
                                }
                                else if (userName == "sandhesh r") //
                                {
                                    userName = "Sandhesh S R";
                                }
                                else if (userName == "S Nandini1")
                                {
                                    userName = "B S Nandini";
                                }
                                //else if (userName == "Sangeetha A")
                                //{
                                //    userName = "Sangeetha Malve";
                                //}
                                else if (userName == "Sulaiman S")
                                {
                                    userName = "Sulaiman Khan S";
                                }
                                else if (userName == "Suprith T")
                                {
                                    userName = "Suprith B T";
                                }
                                else if (userName == "zeeshan1")
                                {
                                    userName = "Zeeshan 1";
                                }

                                if (userName == "Arvindh G")
                                {
                                    userName = "Arvindh A G";
                                }
                                else if (userName == "Sangeetha A")
                                {
                                    userName = "Sangeetha A";
                                }
                                else if (userName == "Maria S")
                                {
                                    userName = "Maria Joseph S";
                                }
                                //else
                                //{
                                //    continue;
                                //}

                                
                                string expression = "agent_name Like '%" + userName.Trim() + "%' and start_date ='" + date + "'";
                                var rows = dt.Select(expression);
                                if (rows.Any())
                                {
                                    dtResult = new DataTable();
                                    dtResult = rows.CopyToDataTable();

                                    Double defaultTime =
                                   dtResult
                                       .AsEnumerable()
                                       .Where(r => (String)r["campaign_name"] == "Default")
                                       .Sum(r => (Int32)r["Total_Time"]);

                                    if (defaultTime > 0)
                                    {
                                        expression = "campaign_name <> 'Default'";
                                        DataTable dtClients = new DataTable();
                                        var clientRows = dtResult.Select(expression);
                                        if (clientRows.Any())
                                        {
                                            dtClients = clientRows.CopyToDataTable();
                                        }
                                        //dtClients = dtResult.Select(expression).CopyToDataTable(); //dtResult.AsEnumerable().Where(r => (String)r["campaign_name"] != "Default");
                                        if (dtClients.Rows.Count > 0)
                                        {
                                            double splitTime = (defaultTime / dtClients.Rows.Count);

                                            dtClients.Rows.Cast<DataRow>().ToList().ForEach(r => r.SetField("Total_Time", UpdateTime(r["Total_Time"].ToString(), splitTime)));


                                            foreach (DataRow finalRow in dtClients.Rows)
                                            {
                                                double RTMDuration = Convert.ToDouble(finalRow["Total_Time"].ToString()) / 60;
                                                TimeSpan span = TimeSpan.FromMinutes(RTMDuration);
                                                string totalDuration = span.ToString(@"hh\:mm\:ss");
                                                string startDateTime = finalRow["start_date"].ToString() + " " + finalRow["start_time"].ToString();
                                                int clientId = 0;
                                               
                                                var query = dtRTMClients.AsEnumerable().Where(x => x.Field<string>("CL_ClientName") == finalRow["campaign_name"].ToString());
                                                foreach (var st in query)
                                                {
                                                    clientId = st.Field<int>("CL_ID");
                                                }

                                                if (clientId != 0)
                                                {
                                                    InsertData(12, empId, user, clientId, 1591, 10239, totalDuration, startDateTime, "Completed", "VCCU", startDateTime);
                                                }
                                              

                                                //Task=1591
                                                //SubTask=10239
                                            }
                                        }
                                        else
                                        {
                                            foreach (DataRow finalRow in dtResult.Rows)
                                            {
                                                double RTMDuration = Convert.ToDouble(finalRow["Total_Time"].ToString()) / 60;
                                                TimeSpan span = TimeSpan.FromMinutes(RTMDuration);
                                                string totalDuration = span.ToString(@"hh\:mm\:ss");
                                                string startDateTime = finalRow["start_date"].ToString() + " " + finalRow["start_time"].ToString();

                                                InsertData(12, empId, user, 18439, 1591, 10239, totalDuration, startDateTime, "Completed", "VCCU", startDateTime);

                                            }
                                        }

                                    }
                                    else
                                    {
                                        foreach (DataRow finalRow in dtResult.Rows)
                                        {
                                            double RTMDuration = Convert.ToDouble(finalRow["Total_Time"].ToString()) / 60;
                                            TimeSpan span = TimeSpan.FromMinutes(RTMDuration);
                                            string totalDuration = span.ToString(@"hh\:mm\:ss");
                                            string startDateTime = finalRow["start_date"].ToString() + " " + finalRow["start_time"].ToString();

                                            int clientId = 0;
                                            //var client = dtRTMClients
                                            //       .AsEnumerable()
                                            //       .Where(r => (String)r["CL_ClientName"] == finalRow["campaign_name"].ToString()).ToArray();
                                            var query = dtRTMClients.AsEnumerable().Where(x => x.Field<string>("CL_ClientName") == finalRow["campaign_name"].ToString());
                                            foreach (var st in query)
                                            {
                                                clientId = st.Field<int>("CL_ID");
                                            }                                            

                                            if (clientId != 0)
                                            {
                                                InsertData(12, empId, user, clientId, 1591, 10239, totalDuration, startDateTime, "Completed", "VCCU", startDateTime);
                                            }
                                            //Task=1591
                                            //SubTask=10239
                                        }
                                    }
                                    //
                                }
                            }
                        }
                    }

                    if (dtAgentState.Rows.Count > 0)
                    {
                        foreach (var selDate in dates)
                        {
                            string date = selDate.ToShortDateString();
                            string startDateTime = date;
                            foreach (DataRow drUser in dtUsers.Rows)
                            {
                                string empId = drUser["UL_Employee_Id"].ToString().Trim();
                                string user = drUser["UL_User_Name"].ToString().Trim();
                                string userName = drUser["UL_System_User_Name"].ToString();
                                userName = userName.Substring(5, userName.Length - 5);
                                userName = userName.Replace('.', ' ');

                                if (userName == "ashwini1")
                                {
                                    userName = "Ashwini Raghunathan";
                                }
                                else if (userName == "c nikitha")
                                {
                                    userName = "Nikitha C";
                                }
                                else if (userName == "Dinesh S")
                                {
                                    userName = "Dinesh Babu S";
                                }
                                else if (userName == "Edward A")
                                {
                                    userName = "Edward Naveen A";
                                }
                                else if (userName == "faizal a") //
                                {
                                    userName = "Faizal Rahman";
                                }
                                else if (userName == "fantin j") //
                                {
                                    userName = "Fantin Cyril";
                                }
                                else if (userName == "sandeep v") //
                                {
                                    userName = "Sandeep Rathnakaran";
                                }
                                else if (userName == "sandhesh r") //
                                {
                                    userName = "Sandhesh S R";
                                }
                                else if (userName == "S Nandini1")
                                {
                                    userName = "B S Nandini";
                                }
                                //else if (userName == "Sangeetha A")
                                //{
                                //    userName = "Sangeetha Malve";
                                //}
                                else if (userName == "Sulaiman S")
                                {
                                    userName = "Sulaiman Khan S";
                                }
                                else if (userName == "Suprith T")
                                {
                                    userName = "Suprith B T";
                                }
                                else if (userName == "zeeshan1")
                                {
                                    userName = "Zeeshan 1";
                                }

                                if (userName == "Arvindh G")
                                {
                                    userName = "Arvindh A G";
                                }
                                else if (userName == "Sangeetha A")
                                {
                                    userName = "Sangeetha A";
                                }
                                else if (userName == "Maria S")
                                {
                                    userName = "Maria Joseph S";
                                }
                                //else
                                //{
                                //    continue;
                                //}
                                

                                string expression = "Agent_Name Like '%" + userName.Trim() + "%' and start_date >= '" + Convert.ToDateTime(date) + "' and start_date < '" + Convert.ToDateTime(date).AddDays(1) + "'";
                                var rows = dtAgentState.Select(expression);
                                if (rows.Any())
                                {
                                    dtResult = new DataTable();
                                    dtResult = rows.CopyToDataTable();

                                    double RTMDuration;
                                    string totalDuration;
                                    double SupportTime = dtResult
                                           .AsEnumerable()
                                           .Where(r => ((r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Unavailable" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Special Project" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Outbound Dialing" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Technical Issue" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Floor Walking" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "HeldPartyAbandoned" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Refused" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Support" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Task" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Disposition" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Offline"))
                                           .Sum(r => (Int32)r["Duration"]);

                                    if (SupportTime > 0)
                                    {
                                        RTMDuration = SupportTime / 60000;
                                        TimeSpan span1 = TimeSpan.FromMinutes(RTMDuration);
                                        totalDuration = span1.ToString(@"hh\:mm\:ss");

                                        InsertData(12, empId, user, 13550, 1326, 10240, totalDuration, startDateTime, "Completed", "VCCU", startDateTime);
                                        //Client = 13550
                                        //Task =1326
                                        //Subtask =10240
                                    }



                                    double TrainingTime = dtResult
                                           .AsEnumerable()
                                           .Where(r => ((r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "New Hire Training" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Meeting" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Up-Training"))
                                           .Sum(r => (Int32)r["Duration"]);

                                    if (TrainingTime > 0)
                                    {
                                        RTMDuration = TrainingTime / 60000;
                                        TimeSpan span2 = TimeSpan.FromMinutes(RTMDuration);
                                        totalDuration = span2.ToString(@"hh\:mm\:ss");

                                        InsertData(12, empId, user, 13550, 1326, 10241, totalDuration, startDateTime, "Completed", "VCCU", startDateTime);
                                        //Client = 13550
                                        //Task =1326
                                        //Subtask =10241
                                    }

                                    double AvailableTime = dtResult
                                           .AsEnumerable()
                                           .Where(r => (r["Skill_Name"] == DBNull.Value ? "" : (String)r["Skill_Name"]) == "" && (Int32)r["Outstate_Code"] == 0)
                                           .Sum(r => (Int32)r["Duration"]);

                                    if (AvailableTime > 0)
                                    {
                                        RTMDuration = AvailableTime / 60000;
                                        TimeSpan span3 = TimeSpan.FromMinutes(RTMDuration);
                                        totalDuration = span3.ToString(@"hh\:mm\:ss");

                                        InsertData(12, empId, user, 13550, 1326, 10242, totalDuration, startDateTime, "Completed", "VCCU", startDateTime);
                                        //Client = 13550
                                        //Task =1326
                                        //Subtask =10242
                                    }
                                }
                                
                            }
                        }
                    }
                }                
                
                label1.Text = "Process Completed Successfully";
                label3.Text = "End time - " + DateTime.Now;
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                label1.Text = "Error occured please try again";
            }
            finally
            {
                
            }
            
        }

        public static Double UpdateTime(string defaultTime, double splitTime)
        {
            Double clientTime = 0;

            if (Convert.ToDouble(defaultTime) > 0)
            {
                clientTime = Convert.ToDouble(defaultTime) + splitTime;
            }

            return clientTime;
        }

        private DataTable GetClients()
        {
            using (SqlDataAdapter da = new SqlDataAdapter("select CL_ID, CL_ClientName, CL_TSheetClient from RTM_Client_List With (NOLOCK) where CL_TeamID = 12", conString))
            {
                da.Fill(dtRTMClients);
            }

            return dtRTMClients;
        }

        private DataTable GetUsers()
        {
            using (SqlDataAdapter da = new SqlDataAdapter("select UL_Employee_Id, UL_User_Name, UL_System_User_Name from RTM_User_List With (NOLOCK) where UL_Team_Id =12 and UL_User_Status =1 order by UL_User_Name", conString))
            {
                da.Fill(dtUsers);
            }

            return dtUsers;
        }

        private bool CheckDuplicates(string _clientNo, string _agentNo, string _callDate, string _callTime)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(conString))
                {
                    string sQuery = "Select VCC_ID from RTM_VCC_CallDetail where VCC_ClientNo='" + _clientNo + "' and VCC_SDAgentNo='"+ _agentNo +"' and VCC_CallDate='"+ _callDate +"' and VCC_CallTime='"+ _callTime +"'";
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.CommandText = sQuery;
                        cmd.Connection = con;
                        cmd.CommandTimeout = int.MaxValue;
                        cmd.CommandType = CommandType.Text;
                        con.Open();
                        SqlDataReader dr = cmd.ExecuteReader();
                        if (dr.HasRows)
                        {
                            con.Close();
                            return true;
                        }
                        con.Close();
                    }
                }
            }
            catch (Exception)
            {
                
            }
            return false;
        }

        private void InsertData(int _teamId, string empId, string userName, int client, int task, int subTask, string duration, string startDateTime, string status, string system, string timeDate)
        {
            using (SqlConnection con = new SqlConnection(conString))
            {
                string sQuery = "Insert into RTM_Records (R_TeamId, R_Employee_Id, R_User_Name, R_Client, R_Task, R_SubTask, R_Duration, R_Start_Date_Time,R_CreatedOn, R_Status, R_System, R_TimeDate ) "+
                                 "Values (" + _teamId + ", '"+ empId +"', '"+ userName +"', "+ client +", "+ task +", "+ subTask +", '"+ duration +"', '"+ startDateTime +"', '"+ DateTime.Now +"', '"+ status+"', '"+ system +"', '"+ startDateTime +"')";

                using (SqlCommand cmd = new SqlCommand(sQuery, con))
                {
                    cmd.CommandTimeout = int.MaxValue;
                    con.Open();
                    cmd.ExecuteNonQuery();
                    con.Close();
                }
            }
        }

        private void RecordVCCTime()
        {
            try
            {
                label1.Text = "Process started Please wait....";
                label2.Text = "Start Time - " + DateTime.Now;
                label3.Text = "";
                com.incontact.login.inSideWS objRep = new com.incontact.login.inSideWS();
                DateTime start = Convert.ToDateTime(DateTime.Now.AddDays(-(int)DateTime.Now.DayOfWeek - 7)).ToUniversalTime();
                DateTime end = Convert.ToDateTime(DateTime.Now.AddDays(-(int)DateTime.Now.DayOfWeek-1)).ToUniversalTime();

                var dates = Enumerable.Range(0, 1 + end.Subtract(start).Days)
                              .Select(offset => start.AddDays(offset))
                              .ToArray();
               
                com.incontact.login.inCredentials objinCred = new com.incontact.login.inCredentials();
                objinCred.busNo = 4593343;
                objinCred.password = "59D4CE34-672B-4915-AA66-FC9D742461AB";

                objinCred.partnerPassword = "";
                objinCred.timeZoneName = "";
                objinCred.DialerDSNName = "";
                objinCred.DialerDSNUserID = "";
                objinCred.DialerDSNPassword = "";

                objRep.inCredentialsValue = objinCred;

                //var myXML = objRep.DataDownloadReport_Run(16, start, end);
                var myXML = objRep.DataDownloadReport_Run(16, start, end);
                DataTable dt = new DataTable();
                dt = myXML.Tables[0];

                var agentStateXML = objRep.DataDownloadReport_Run(350083, start, end);

                dtAgentState = agentStateXML.Tables[0];

                dtUsers = GetUsers();
                dtRTMClients = GetClients();
                DataTable dtResult = new DataTable();
                if (dt.Rows.Count > 0)
                {
                        if (dtUsers.Rows.Count > 0)
                        {  
                            string date = DateTime.Now.AddDays(-1).ToShortDateString();
                            foreach (DataRow drUser in dtUsers.Rows)
                            {
                                string empId = drUser["UL_Employee_Id"].ToString().Trim();
                                string user = drUser["UL_User_Name"].ToString().Trim();
                                string userName = drUser["UL_System_User_Name"].ToString();
                                userName = userName.Substring(5, userName.Length - 5);
                                userName = userName.Replace('.', ' ');
                                
                                string expression = "agent_name Like '%" + userName.Trim() + "%' and start_date ='" + date + "'";
                                var rows = dt.Select(expression);
                                if (rows.Any())
                                {
                                    dtResult = new DataTable();
                                    dtResult = rows.CopyToDataTable();

                                    Double defaultTime =
                                   dtResult
                                       .AsEnumerable()
                                       .Where(r => (String)r["campaign_name"] == "Default" && r["agent_name"].ToString().Contains(userName.Trim()))
                                       .Sum(r => (Int32)r["Total_Time"]);

                                    if (defaultTime > 0)
                                    {
                                        expression = "campaign_name <> 'Default'";
                                        DataTable dtClients = new DataTable();
                                        var clientRows = dtResult.Select(expression);
                                        if (clientRows.Any())
                                        {
                                            dtClients = clientRows.CopyToDataTable();
                                        }
                                        //dtClients = dtResult.Select(expression).CopyToDataTable(); //dtResult.AsEnumerable().Where(r => (String)r["campaign_name"] != "Default");
                                        if (dtClients.Rows.Count > 0)
                                        {
                                            double splitTime = (defaultTime / dtClients.Rows.Count);

                                            dtClients.Rows.Cast<DataRow>().ToList().ForEach(r => r.SetField("Total_Time", UpdateTime(r["Total_Time"].ToString(), splitTime)));


                                            foreach (DataRow finalRow in dtClients.Rows)
                                            {
                                                double RTMDuration = Convert.ToDouble(finalRow["Total_Time"].ToString()) / 60;
                                                TimeSpan span = TimeSpan.FromMinutes(RTMDuration);
                                                string totalDuration = span.ToString(@"hh\:mm\:ss");
                                                string startDateTime = finalRow["start_date"].ToString() + " " + finalRow["start_time"].ToString();
                                                int clientId = 0;
                                               
                                                var query = dtRTMClients.AsEnumerable().Where(x => x.Field<string>("CL_ClientName") == finalRow["campaign_name"].ToString());
                                                foreach (var st in query)
                                                {
                                                    clientId = st.Field<int>("CL_ID");
                                                }

                                                if (clientId != 0)
                                                {
                                                    InsertData(12, empId, user, clientId, 1591, 10239, totalDuration, startDateTime, "Completed", "VCCU", startDateTime);
                                                }
                                              

                                                //Task=1591
                                                //SubTask=10239
                                            }
                                        }
                                        else
                                        {
                                            foreach (DataRow finalRow in dtResult.Rows)
                                            {
                                                double RTMDuration = Convert.ToDouble(finalRow["Total_Time"].ToString()) / 60;
                                                TimeSpan span = TimeSpan.FromMinutes(RTMDuration);
                                                string totalDuration = span.ToString(@"hh\:mm\:ss");
                                                string startDateTime = finalRow["start_date"].ToString() + " " + finalRow["start_time"].ToString();

                                                InsertData(12, empId, user, 18439, 1591, 10239, totalDuration, startDateTime, "Completed", "VCCU", startDateTime);

                                            }
                                        }

                                    }
                                    else
                                    {
                                        foreach (DataRow finalRow in dtResult.Rows)
                                        {
                                            double RTMDuration = Convert.ToDouble(finalRow["Total_Time"].ToString()) / 60;
                                            TimeSpan span = TimeSpan.FromMinutes(RTMDuration);
                                            string totalDuration = span.ToString(@"hh\:mm\:ss");
                                            string startDateTime = finalRow["start_date"].ToString() + " " + finalRow["start_time"].ToString();

                                            int clientId = 0;
                                            //var client = dtRTMClients
                                            //       .AsEnumerable()
                                            //       .Where(r => (String)r["CL_ClientName"] == finalRow["campaign_name"].ToString()).ToArray();
                                            var query = dtRTMClients.AsEnumerable().Where(x => x.Field<string>("CL_ClientName") == finalRow["campaign_name"].ToString());
                                            foreach (var st in query)
                                            {
                                                clientId = st.Field<int>("CL_ID");
                                            }

                                            if (clientId != 0)
                                            {
                                                InsertData(12, empId, user, clientId, 1591, 10239, totalDuration, startDateTime, "Completed", "VCCU", startDateTime);
                                            }
                                            //Task=1591
                                            //SubTask=10239
                                        }
                                    }
                                    //
                                }
                            }
                        
                    }
                }

                    if (dtAgentState.Rows.Count > 0)
                    {
                        
                            string date = DateTime.Now.AddDays(-1).ToShortDateString();
                            string startDateTime = date;
                            foreach (DataRow drUser in dtUsers.Rows)
                            {
                                string empId = drUser["UL_Employee_Id"].ToString().Trim();
                                string user = drUser["UL_User_Name"].ToString().Trim();
                                string userName = drUser["UL_System_User_Name"].ToString();
                                userName = userName.Substring(5, userName.Length - 5);
                                userName = userName.Replace('.', ' ');
                                string expression = "Agent_Name Like '%" + userName.Trim() + "%' and start_date >= '" + Convert.ToDateTime(date) + "' and start_date < '" + Convert.ToDateTime(date).AddDays(1) + "'";
                                var rows = dtAgentState.Select(expression);
                                if (rows.Any())
                                {
                                    dtResult = new DataTable();
                                    dtResult = rows.CopyToDataTable();

                                    double RTMDuration;
                                    string totalDuration;
                                    double SupportTime = dtResult
                                           .AsEnumerable()
                                           .Where(r => ((r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Unavailable" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Outbound Dialing" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Technical Issue" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Floor Walking" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "HeldPartyAbandoned" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Refused" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Disposition" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Offline") && r["Agent_Name"].ToString().Contains(userName.Trim()))
                                           .Sum(r => (Int32)r["Duration"]);

                                    if (SupportTime > 0)
                                    {
                                        RTMDuration = SupportTime / 60000;
                                        TimeSpan span1 = TimeSpan.FromMinutes(RTMDuration);
                                        totalDuration = span1.ToString(@"hh\:mm\:ss");

                                        InsertData(12, empId, user, 13550, 1326, 10240, totalDuration, startDateTime, "Completed", "VCCU", startDateTime);
                                        //Client = 13550
                                        //Task =1326
                                        //Subtask =10240
                                    }



                                    double TrainingTime = dtResult
                                           .AsEnumerable()
                                           .Where(r => ((r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "New Hire Training" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Meeting" || (r["Outstate"] == DBNull.Value ? "" : (String)r["Outstate"]) == "Up-Training") && r["Agent_Name"].ToString().Contains(userName.Trim()))
                                           .Sum(r => (Int32)r["Duration"]);

                                    if (TrainingTime > 0)
                                    {
                                        RTMDuration = TrainingTime / 60000;
                                        TimeSpan span2 = TimeSpan.FromMinutes(RTMDuration);
                                        totalDuration = span2.ToString(@"hh\:mm\:ss");

                                        InsertData(12, empId, user, 13550, 1326, 10241, totalDuration, startDateTime, "Completed", "VCCU", startDateTime);
                                        //Client = 13550
                                        //Task =1326
                                        //Subtask =10241
                                    }

                                    double AvailableTime = dtResult
                                           .AsEnumerable()
                                           .Where(r => (r["Skill_Name"] == DBNull.Value ? "" : (String)r["Skill_Name"]) == "" && (Int32)r["Outstate_Code"] == 0 && r["Agent_Name"].ToString().Contains(userName.Trim()))
                                           .Sum(r => (Int32)r["Duration"]);

                                    if (AvailableTime > 0)
                                    {
                                        RTMDuration = AvailableTime / 60000;
                                        TimeSpan span3 = TimeSpan.FromMinutes(RTMDuration);
                                        totalDuration = span3.ToString(@"hh\:mm\:ss");

                                        InsertData(12, empId, user, 13550, 1326, 10242, totalDuration, startDateTime, "Completed", "VCCU", startDateTime);
                                        //Client = 13550
                                        //Task =1326
                                        //Subtask =10242
                                    }
                                }
                                
                            }
                        
                    }
                             
                
                label1.Text = "Process Completed Successfully";
                label3.Text = "End time - " + DateTime.Now;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                label1.Text = "Error occured please try again";
            }
            finally
            {
                
            }
        }

      
    }
}
