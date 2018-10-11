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


namespace RTMReportsAssistant
{
    public partial class Form1 : Form
    {
        //  SqlConnection globalCon = new SqlConnection(@"Data Source=10.55.5.40,1433;Initial Catalog=RTM_Global_Test;User ID=PRODRTMDB;Password=Prodrtm@123;");

        SqlConnection con = new SqlConnection(@"Data Source=10.55.5.40,1433;Initial Catalog=RTM_Global_Test;User ID=PRODRTMDB;Password=Prodrtm@123;");
        //SqlConnection con = new SqlConnection(@"Data Source=10.55.5.40,1433;Initial Catalog=Real_Time_Metrics_Dev;User ID=PRODRTMDB;Password=Prodrtm@123;");
        SqlConnection globalCon = new SqlConnection(@"Data Source=10.55.5.40,1433;Initial Catalog=RTM_Global_Test;User ID=PRODRTMDB;Password=Prodrtm@123;");
        //SqlConnection globalCon = new SqlConnection(@"Data Source=BLRPRODRTM\RTM_PROD_BLR;Initial Catalog=Real_Time_Metrics;User ID=sa;Password=Prodrtm@123;");
        DataSet ds = new DataSet();
        SqlDataAdapter da;
        SqlCommand cmd;
        StringBuilder myBuilder = new StringBuilder();
        int flag = 0;
        int delayFlag = 0;
        int flagEst = 0;
        int flagEstOnboard = 0;
        int flagEstComp = 0;
        int flagCheckESTUsers = 0;
        int flagOpsCheck = 0;
        int peerFlag = 0;
        int earlyLogOffFlag = 0;
        int flagQCCheck = 0;
        int flagAuditCheck = 0;
        int flagLateLogin = 0;
        int flagAuditEst = 0;
        int flagCS = 0;
        int HRISFlag = 0;
        int flagOffshore = 0;
        int flagIncompleteData = 0;
        int flagSubmitStatus = 0;
        int flagNonComplainceReport = 0;
        int weeklyHoursFlag = 0;
        int flagNonComplainceReportAsentinal = 0;
        int oirFlag = 0;
        int RPAfalg = 0;
        int EffectiveRateReportDaily = 0;
        int EffectiveRateReportMonthly = 0;


        DataTable dtResult = new DataTable();
        DataTable dt = new DataTable();
        DataTable dtInst = new DataTable();
        string FromAddress = System.Configuration.ConfigurationManager.AppSettings["FromAddress"].ToString();
        // int port = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["FromAddress"]);
        string host = System.Configuration.ConfigurationManager.AppSettings["SMTPClient"].ToString();
        string[] ToAddrSuccess = System.Configuration.ConfigurationManager.AppSettings["ToAddressSuccess"].Split(',');
        string[] OnBoardToAddress = System.Configuration.ConfigurationManager.AppSettings["OnBoardToAddresses"].Split(',');
        string[] cs1ToAddress = System.Configuration.ConfigurationManager.AppSettings["CS1"].Split(',');
        string[] cs2ToAddress = System.Configuration.ConfigurationManager.AppSettings["CS2"].Split(',');
        private static TimeZoneInfo INDIAN_ZONE = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SendPSLDailyER_QC();
           //To run scheduler make Enabled=true
            tmrUpdate.Enabled = true;
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
                //this.notifyIcon1.ShowBalloonTip(5000);
                this.ShowInTaskbar = false;
            }
        }

        private void tmrUpdate_Tick(object sender, EventArgs e)
        {

            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
           
            try
            {               

                if (indianTime.DayOfWeek == DayOfWeek.Monday && indianTime.Hour == 3)
                {
                    if (flag == 0)
                    {
                        flag = 1;
                        //loadRUDetails();
                    }
                }

                if (indianTime.DayOfWeek == DayOfWeek.Saturday && indianTime.Hour == 19)
                {
                    if (RPAfalg == 0)
                    {
                        RPAfalg = 1;
                        GenerateRPAReport();
                    }
                }
                

                if (indianTime.DayOfWeek == DayOfWeek.Tuesday && indianTime.Hour == 6)
                {
                    if (flagSubmitStatus == 0)
                    {
                        flagSubmitStatus = 1;
                        SendSubmitStatusToRashmi();
                    }
                }

                if (indianTime.DayOfWeek == DayOfWeek.Wednesday && indianTime.Hour == 9)
                {
                    if (weeklyHoursFlag == 0)
                    {
                        weeklyHoursFlag = 1;
                        WeeklyTrackingReport();
                    }
                }

                if (indianTime.DayOfWeek == DayOfWeek.Thursday && indianTime.Hour == 6)
                {
                    weeklyHoursFlag = 0;
                    flagSubmitStatus = 0;
                    RPAfalg = 0;
                }

                //Effective Rate Report for Invoices Team: weekly Report*
                if (indianTime.DayOfWeek == DayOfWeek.Monday && indianTime.Hour == 6)//6
                {
                    if (earlyLogOffFlag == 0)
                    {
                        earlyLogOffFlag = 1;
                        SendEarlyLogOffDetails();
                        //Effective Rate Report : weekly Report***Namohar Modified 17/08/2018.
                        SendWeeklySKUEffectiveRateFromSKUDB();
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToErrorLog(ex.Message, ex.StackTrace, "ResourceUtilizationError");
            }

            //**************   //Effective Rate Report for Invoices Team: Daily Report* ***Namohar Modified 17/08/2018.
            try
            {
                if (indianTime.Hour == 6)
                {
                    if (EffectiveRateReportDaily == 0)
                    {
                        EffectiveRateReportDaily = 1;
                       // Daily Report
                        SendWeeklySKUEffectiveRateFromSKUDBDay();
                    }
                }
                if (indianTime.Hour == 9)
                {
                    EffectiveRateReportDaily = 0;
                }
            }
            catch (Exception ex)
            {
                WriteToErrorLog(ex.Message, ex.StackTrace, "PeerSupportError");
            }

            //Effective Rate Report for Invoices Team: Monthly Report* ***Namohar Modified 17/08/2018.
            try
            {

                DateTime now = DateTime.Now;
                DateTime firstDay = new DateTime(now.Year, now.Month, 1);
                if (firstDay.Day == indianTime.Day)
                {
                    if (indianTime.Hour == 6)
                    {
                        if (EffectiveRateReportMonthly == 0)
                        {
                            EffectiveRateReportMonthly = 1;
                            // Monthly Report*                       
                          SendWeeklySKUEffectiveRateFromSKUDBMonthly();
                        }
                    }
                    if (indianTime.Hour == 7)
                    {
                        EffectiveRateReportMonthly = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToErrorLog(ex.Message, ex.StackTrace, "PeerSupportError");
            }



            try
            {
                if (indianTime.DayOfWeek == DayOfWeek.Monday && indianTime.Hour == 5)//5
                {
                    if (peerFlag == 0)
                    {
                        peerFlag = 1;
                        //WeeklyPeerSupport();
                    }
                }

                if (indianTime.DayOfWeek == DayOfWeek.Monday && indianTime.Hour == 11)//5
                {
                    if (flagOffshore == 0)
                    {
                        flagOffshore = 1;
                        OffshoreTasks();
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToErrorLog(ex.Message, ex.StackTrace, "PeerSupportError");
            }

            try
            {
                if (indianTime.DayOfWeek == DayOfWeek.Monday && indianTime.Hour == 5) //5
                {
                    if (delayFlag == 0)
                    {
                        delayFlag = 1;
                        //Delay Logins
                        BindDelayDetails();
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToErrorLog(ex.Message, ex.StackTrace, "DelayDetailsError");
            }

            try
            {
                if (indianTime.Hour == 15)
                {
                    if (flagEstOnboard == 0)
                    {
                        flagEstOnboard = 1;
                        //DisplayEstRecords(10);
                    }
                }
                else if (indianTime.Hour == 14)
                {
                    if (flagEst == 0)
                    {
                        flagEst = 1;
                        //DisplayEstRecords(13);
                        //DisplayEstRecords(31);
                    }
                }
                else if (indianTime.Hour == 11)
                {
                    if (flagEst == 0)
                    {
                        flagEst = 1;
                        //DisplayEstRecords(18);

                        //SendSKUEffectiveRateFromSKUDB();                   
                       SendCMPEffectiveRateFromCMPDB();
                        //PSLDB*****
                        SendPSLDailyER_QC();

                    }
                }

                if (indianTime.Hour == 15 || indianTime.Hour == 12)
                {
                    flagEst = 0;
                }

                if (indianTime.Hour == 16)
                {
                    flagEstOnboard = 0;
                }

                //if (DateTime.Now.Hour == 12)
                //{
                //    if (flagQCEst == 0)
                //    {
                //        flagQCEst = 1;
                //        DisplayEstRecords(9);
                //    }
                //}

                if (indianTime.Hour == 12)
                {
                    if (flagCS == 0)
                    {
                        flagCS = 1;
                        //DisplayEstRecords(22);
                        DisplayEstRecords(23);
                        //DisplayEstRecords(29);
                    }
                }
                if (indianTime.Hour == 13)
                {
                    flagCS = 0;
                }

                //if (DateTime.Now.Hour == 11 && DateTime.Now.Minute == 30)
                //{
                //    if (flagAuditEst == 0)
                //    {
                //        flagAuditEst = 1;
                //        DisplayEstRecords(11);
                //    }
                //}

                //if (DateTime.Now.Hour == 13)
                //{
                //    flagQCEst = 0;
                //    flagAuditEst = 0;
                //}
            }
            catch (Exception ex)
            {
                WriteToErrorLog(ex.Message, ex.StackTrace, "EstimationError");
            }

            try
            {

               
             
                if (indianTime.Hour == 10)
                {
                    if (flagEstComp == 0)
                    {
                        flagEstComp = 1;
                        //DisplayEstCompData(10);
                        //IPV -invoices.
                        SendSampleSKUEffectiveRateFromSKUDB();

                        //CMPDB-QC.
                       SendCMPDailyER_QC();

                        //IPV - QC.
                        SendSKUDailyER_QC();
                        //Inventory.
                        ER_IN_Inventory_Management();
                        //Not firing , need to check
                        ER_IN_Provisioning();

                      
                    }
                }
                else if (indianTime.Hour == 8)
                {
                    if (flagEstComp == 0)
                    {
                        flagEstComp = 1;
                        //DisplayEstCompData(13);
                        //DisplayEstCompData(18);
                        //DisplayEstCompData(9);
                        //DisplayEstCompData(11);
                        // DisplayEstCompData(22);
                        DisplayEstCompData(23);
                        //DisplayEstCompData(29);
                        //DisplayEstCompData(31);

                    }
                }
                if (indianTime.Hour == 9 || indianTime.Hour == 11)
                {
                    flagEstComp = 0;
                }
            }
            catch (Exception ex)
            {
                WriteToErrorLog(ex.Message, ex.StackTrace, "EstimationComparisonError");

            }
            try
            {
                if (indianTime.Hour == 14)
                {

                    if (flagCheckESTUsers == 0)
                    {
                        flagCheckESTUsers = 1;
                        // CheckEstUsers(10);
                    }

                }
                else if (indianTime.Hour == 12)
                {
                    if (flagCheckESTUsers == 0)
                    {
                        flagCheckESTUsers = 1;
                        //CheckEstUsers(13);
                        //CheckEstUsers(31);
                    }
                }

                if (indianTime.Hour == 11)
                {
                    if (flagQCCheck == 0)
                    {
                        flagQCCheck = 1;
                        // CheckEstUsers(9);
                    }
                }
                if (indianTime.Hour == 9)
                {
                    if (flagOpsCheck == 0)
                    {
                        flagOpsCheck = 1;
                        //CheckEstUsers(18);
                    }
                }
                if (indianTime.Hour == 11 && indianTime.Minute == 30)
                {
                    if (flagAuditCheck == 0)
                    {
                        flagAuditCheck = 1;
                        // CheckEstUsers(11);
                    }
                }

                if (indianTime.Hour == 10 && indianTime.Minute == 30)
                {
                    if (flagCS == 0)
                    {
                        flagCS = 1;
                        // CheckEstUsers(22);
                        CheckEstUsers(23);
                        //CheckEstUsers(29);
                    }
                }
                if (indianTime.Hour == 11)
                {
                    flagCS = 0;
                }


                if (indianTime.Hour == 12)
                {
                    flagQCCheck = 0;
                    flagAuditCheck = 0;

                }

                if (indianTime.Hour == 13 || indianTime.Hour == 16)
                {
                    flagCheckESTUsers = 0;
                }
                if (DateTime.Now.Hour == 10)
                {
                    flagOpsCheck = 0;
                }
            }
            catch (Exception ex)
            {
                WriteToErrorLog(ex.Message, ex.StackTrace, "EstimationError");
            }


            if (indianTime.DayOfWeek == DayOfWeek.Tuesday)
            {
                flag = 0;
                delayFlag = 0;
                peerFlag = 0;
                earlyLogOffFlag = 0;
                flagOffshore = 0;
                //flagIncompleteData = 0;
            }

            //if (indianTime.DayOfWeek == DayOfWeek.Wednesday)
            //{
            //    flagIncompleteData = 0;
            //    flagNonComplainceReport = 0;
            //}
            //loadRUDetails();

            //Late login notifications
            try
            {
                if (indianTime.Hour == 8 && indianTime.Minute == 30)
                {
                    if (flagLateLogin == 0)
                    {
                        flagLateLogin = 1;
                        SendLateLoginDetails(DateTime.Now.ToShortDateString(), 6, "07:00", "08:30");//WEM
                        //SendLateLoginDetails(DateTime.Now.ToShortDateString(), 7, "07:00", "08:30");//Invoice3
                        //SendLateLoginDetails(DateTime.Now.ToShortDateString(), 8, "07:00", "08:30");//Invoice2
                        SendLateLoginDetails(DateTime.Now.ToShortDateString(), 29, "07:00", "08:30"); //Invoices
                        SendLateLoginDetails(DateTime.Now.ToShortDateString(), 11, "07:00", "08:30");//Audit
                        SendLateLoginDetails(DateTime.Now.ToShortDateString(), 13, "07:00", "08:30");//Implementation
                        //SendLateLoginDetails(DateTime.Now.ToShortDateString(), 14, "07:00", "08:30");//Catelog
                        SendLateLoginDetails(DateTime.Now.ToShortDateString(), 18, "07:00", "08:30");//Ops Support
                    }

                }

                if (indianTime.Hour == 10)
                {
                    if (flagLateLogin == 0)
                    {
                        flagLateLogin = 1;
                        SendLateLoginDetails(DateTime.Now.ToShortDateString(), 1, "00:00", "09:55");
                        SendLateLoginDetails(DateTime.Now.ToShortDateString(), 9, "08:00", "10:00"); // QC
                        SendLateLoginDetails(DateTime.Now.ToShortDateString(), 0, "00:00", "09:55");//Managers
                    }
                }

                if (indianTime.Hour == 15)
                {
                    if (flagLateLogin == 0)
                    {
                        flagLateLogin = 1;
                        SendLateLoginDetails(DateTime.Now.ToShortDateString(), 1, "09:56", "14:55");
                        SendLateLoginDetails(DateTime.Now.ToShortDateString(), 10, "09:56", "14:55");//Onboarding
                        SendLateLoginDetails(DateTime.Now.ToShortDateString(), 0, "09:56", "14:55");//Managers
                        SendLateLoginDetails(DateTime.Now.ToShortDateString(), 16, "07:30", "14:55"); //IN_Corelogic
                        SendLateLoginDetails(DateTime.Now.ToShortDateString(), 120, "07:30", "14:55"); //IN_Admin
                    }
                }

                if (indianTime.Hour == 20)
                {
                    if (flagLateLogin == 0)
                    {
                        flagLateLogin = 1;
                        SendLateLoginDetails(DateTime.Now.ToShortDateString(), 1, "14:56", "19:55");
                        SendLateLoginDetails(DateTime.Now.ToShortDateString(), 0, "14:56", "19:55");//Managers
                    }
                }

                if (indianTime.Hour == 23 && indianTime.Minute == 59)
                {
                    if (flagLateLogin == 0)
                    {
                        flagLateLogin = 1;
                        SendLateLoginDetails(DateTime.Now.ToShortDateString(), 1, "19:56", "23:59");
                        SendLateLoginDetails(DateTime.Now.ToShortDateString(), 0, "19:56", "23:59");//Managers
                    }
                }

                if (indianTime.Hour == 21)
                {
                    if (oirFlag == 0)
                    {
                        oirFlag = 1;
                        Invoices_OIR_Processing();
                    }
                }

                if (indianTime.Hour == 9 || indianTime.Hour == 11 || indianTime.Hour == 16 || indianTime.Hour == 21 || indianTime.Hour == 7)
                {
                    flagLateLogin = 0;
                }

                if (indianTime.Hour == 23)
                {
                    oirFlag = 0;
                }
            }
            catch (Exception ex)
            {
                WriteToErrorLog(ex.Message, ex.StackTrace, "LateLogin");
            }


            //Comment end here

            // RTM data imcomplete reminder

            try
            {
                if (DateTime.Now.DayOfWeek == DayOfWeek.Tuesday && DateTime.Now.Hour == 6 && DateTime.Now.Minute == 30)
                {
                    if (flagIncompleteData == 0)
                    {
                        flagIncompleteData = 1;
                        incompleteData();
                    }
                }

                if (DateTime.Now.DayOfWeek == DayOfWeek.Wednesday)
                {
                    flagIncompleteData = 0;
                    flagNonComplainceReport = 0;
                    flagNonComplainceReportAsentinal = 0;
                }

            }
            catch (Exception ex)
            {
                WriteToErrorLog(ex.Message, ex.StackTrace, "RTM Reminder error");
            }


            try
            {
                if (DateTime.Now.DayOfWeek == DayOfWeek.Tuesday && DateTime.Now.Hour == 12 && DateTime.Now.Minute == 30)
                {
                    if (flagNonComplainceReport == 0)
                    {
                        flagNonComplainceReport = 1;
                        NonComplainceReport();
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToErrorLog(ex.Message, ex.StackTrace, "Non Complaince Report error");
            }


            try
            {

                if (DateTime.Now.DayOfWeek == DayOfWeek.Tuesday && DateTime.Now.Hour == 13 && DateTime.Now.Minute == 30)
                {
                    if (flagNonComplainceReportAsentinal == 0)
                    {
                        flagNonComplainceReportAsentinal = 1;
                        NonComplainceReportAsentinel();
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToErrorLog(ex.Message, ex.StackTrace, "Non Complaince Report error");
            }
        }

        private void BuildLateLoginTable()
        {
            dtResult = new DataTable();
            DataColumn dc;

            dc = new DataColumn("User Name");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Employee Id");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Date");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Sceduled Login");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("First Activity");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Delayed Login");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Delay Reason");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Total delay instances in this month");
            dtResult.Columns.Add(dc);
        }

        private void SendLateLoginDetails(string date, int teamid, string fromTime, string toTime)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            var first = new DateTime(indianTime.Year, indianTime.Month, 1);
            dt = new DataTable();
            DataRow dr;
            BuildLateLoginTable();
            if (teamid == 29)
            {
                da = new SqlDataAdapter("select D_UserName as [User Name],UL_Employee_Id as [Employee Id], CONVERT(VARCHAR(10), D_Date, 101) as [Date],CONVERT(VARCHAR(10), D_SLogin, 108) as [Sceduled Login], CONVERT(VARCHAR(10), D_Date, 108) as [First Activity], D_Duration as [Delayed Login], D_Reason as [Delay Reason] " +
                    "from dbo.RTM_DelayedLogInOff, dbo.RTM_User_List where D_UserName = UL_User_Name and SUBSTRING( convert(varchar, D_Date,108),1,5) between '" + fromTime + "' and '" + toTime + "' " +
                     "and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, D_Date))) = '" + date + "' and (D_Team_Id = 29 or D_Team_Id = 14) and D_Type = 'In' Order By D_Reason, D_UserName", con);
            }
            else if (teamid == 0)
            {
                da = new SqlDataAdapter("select D_UserName as [User Name],UL_Employee_Id as [Employee Id], CONVERT(VARCHAR(10), D_Date, 101) as [Date],CONVERT(VARCHAR(10), D_SLogin, 108) as [Sceduled Login], CONVERT(VARCHAR(10), D_Date, 108) as [First Activity], D_Duration as [Delayed Login], D_Reason as [Delay Reason] " +
                    "from dbo.RTM_DelayedLogInOff, dbo.RTM_User_List where D_UserName = UL_User_Name and SUBSTRING( convert(varchar, D_Date,108),1,5) between '" + fromTime + "' and '" + toTime + "' " +
                     "and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, D_Date))) = '" + date + "' and UL_RepMgrId='102651' and D_Type = 'In' Order By D_Reason, D_UserName", con);
            }
            else
            {
                da = new SqlDataAdapter("select D_UserName as [User Name],UL_Employee_Id as [Employee Id], CONVERT(VARCHAR(10), D_Date, 101) as [Date],CONVERT(VARCHAR(10), D_SLogin, 108) as [Sceduled Login], CONVERT(VARCHAR(10), D_Date, 108) as [First Activity], D_Duration as [Delayed Login], D_Reason as [Delay Reason] " +
                    "from dbo.RTM_DelayedLogInOff, dbo.RTM_User_List where D_UserName = UL_User_Name and SUBSTRING( convert(varchar, D_Date,108),1,5) between '" + fromTime + "' and '" + toTime + "' " +
                     "and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, D_Date))) = '" + date + "' and D_Team_Id = " + teamid + " and D_Type = 'In' Order By D_Reason, D_UserName", con);
            }
            da.Fill(dt);

            TimeSpan totalLateDuration = TimeSpan.Parse("00:00:00");
            int totalInstances = 0;
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr1 in dt.Rows)
                {
                    dr = dtResult.NewRow();

                    dr["User Name"] = dr1["User Name"];
                    dr["Employee Id"] = dr1["Employee Id"];
                    dr["Date"] = dr1["Date"];
                    dr["Sceduled Login"] = dr1["Sceduled Login"];
                    dr["First Activity"] = dr1["First Activity"];
                    dr["Delayed Login"] = dr1["Delayed Login"];
                    totalLateDuration = totalLateDuration.Add(TimeSpan.Parse(dr1["Delayed Login"].ToString()));
                    dr["Delay Reason"] = dr1["Delay Reason"];
                    dtInst = new DataTable();
                    using (da = new SqlDataAdapter("select D_UserName, convert(varchar(10),D_Date,101) as [Date] from RTM_DelayedLogInOff where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, D_Date))) between '" + Convert.ToDateTime(first).ToShortDateString() + "' and '" + indianTime.ToShortDateString() + "' and D_UserName ='" + dr1["User Name"].ToString() + "' and D_Type = 'In' group by D_UserName, convert(varchar(10),D_Date,101)", con))
                    {
                        da.Fill(dtInst);
                    }
                    if (dtInst.Rows.Count > 0)
                    {
                        dr["Total delay instances in this month"] = dtInst.Rows.Count.ToString();
                        totalInstances = totalInstances + dtInst.Rows.Count;
                    }
                    else
                    {
                        dr["Total delay instances in this month"] = "0";
                    }
                    dtResult.Rows.Add(dr);
                }

                dr = dtResult.NewRow();

                dr["User Name"] = "";
                dr["Employee Id"] = "";
                dr["Date"] = "";
                dr["Sceduled Login"] = "";
                dr["First Activity"] = "Total";
                dr["Delayed Login"] = totalLateDuration;

                dr["Delay Reason"] = "";
                dr["Total delay instances in this month"] = totalInstances;
                dtResult.Rows.Add(dr);

                getLateLoginHTML(dtResult);

                StringBuilder sb = new StringBuilder();

                sb.AppendLine("");
                sb.AppendLine(myBuilder.ToString());
                sb.AppendLine("");//here I want the data to       display in table format
                sb.AppendLine("");
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                sb.AppendLine("");

                MailMessage message1 = new MailMessage();
                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);

                if (teamid == 0)
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Managers -  Late Login Summary (" + indianTime.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }
                if (teamid == 1) // OrderDesk
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Sandeep.C@tangoe.com"));
                    message1.To.Add(new MailAddress("Suresh.Subbarathaniah@tangoe.com"));
                    message1.To.Add(new MailAddress("Vicky.Ghodke@tangoe.com"));
                    message1.To.Add(new MailAddress("E.Sumanth@tangoe.com"));
                    message1.To.Add(new MailAddress("Rajesh.Subramanyam@tangoe.com"));
                    message1.To.Add(new MailAddress("Shabeenaz1@tangoe.com"));
                    message1.To.Add(new MailAddress("Muralidharan.Parthasarathy@tangoe.com"));
                    message1.To.Add(new MailAddress("Sumit.Bhat@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Fulfillment Team -  Late Login Summary (" + indianTime.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 6) //Wem
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Arjun.Nagraj@tangoe.com"));
                    message1.To.Add(new MailAddress("Rajesh.Subramanyam@tangoe.com"));
                    message1.To.Add(new MailAddress("Shabeenaz1@tangoe.com"));
                    message1.To.Add(new MailAddress("Muralidharan.Parthasarathy@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "WEM Team -  Late Login Summary (" + indianTime.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 7) // Invoice 3 -------------------------------------
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    //message1.To.Add(new MailAddress("Balaji.Nagabhushanam@tangoe.com"));
                    message1.To.Add(new MailAddress("Sriram.Krishnan@tangoe.com"));
                    message1.To.Add(new MailAddress("Vikas.Vyas@tangoe.com"));
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Rajesh.Subramanyam@tangoe.com"));
                    message1.To.Add(new MailAddress("Shabeenaz1@tangoe.com"));
                    message1.To.Add(new MailAddress("Muralidharan.Parthasarathy@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Invoice 3 Team -  Late Login Summary (" + indianTime.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 29) //Invoice 2
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Vikas.Vyas@tangoe.com"));
                    message1.To.Add(new MailAddress("Sriram.Krishnan@tangoe.com"));
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Rajesh.Subramanyam@tangoe.com"));
                    message1.To.Add(new MailAddress("Shabeenaz1@tangoe.com"));
                    message1.To.Add(new MailAddress("Muralidharan.Parthasarathy@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Invoices -  Late Login Summary (" + indianTime.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 9) // Quality Check
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Rajesh.Subramanyam@tangoe.com"));
                    message1.To.Add(new MailAddress("Shabeenaz1@tangoe.com"));
                    message1.To.Add(new MailAddress("Muralidharan.Parthasarathy@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Quality Check Team -  Late Login Summary (" + indianTime.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 10) //Onboarding
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Balaji.Nagabhushanam@tangoe.com"));
                    message1.To.Add(new MailAddress("Rajesh.Subramanyam@tangoe.com"));
                    message1.To.Add(new MailAddress("Shabeenaz1@tangoe.com"));
                    message1.To.Add(new MailAddress("Muralidharan.Parthasarathy@tangoe.com"));
                    //message1.To.Add(new MailAddress("Sumit.Bhat@tangoe.com"));
                    message1.To.Add(new MailAddress("Harish.Sadananda@tangoe.com"));
                    message1.To.Add(new MailAddress("Piyali.Bhattacharjee@tangoe.com"));
                    message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Onboarding Team -  Late Login Summary (" + indianTime.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 11) //Audit & Optimize
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Vinith.Bekal@tangoe.com"));
                    message1.To.Add(new MailAddress("Rajesh.Subramanyam@tangoe.com"));
                    message1.To.Add(new MailAddress("Shabeenaz1@tangoe.com"));
                    message1.To.Add(new MailAddress("Muralidharan.Parthasarathy@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Audit & Optimize Team -  Late Login Summary (" + indianTime.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 13) // Implementation
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Balaji.Nagabhushanam@tangoe.com"));
                    message1.To.Add(new MailAddress("Rajesh.Subramanyam@tangoe.com"));
                    message1.To.Add(new MailAddress("Shabeenaz1@tangoe.com"));
                    message1.To.Add(new MailAddress("Muralidharan.Parthasarathy@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Implementation Team -  Late Login Summary (" + indianTime.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 14) //Catalog & Mapping
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Vikas.Vyas@tangoe.com"));
                    message1.To.Add(new MailAddress("Rajesh.Subramanyam@tangoe.com"));
                    message1.To.Add(new MailAddress("Shabeenaz1@tangoe.com"));
                    message1.To.Add(new MailAddress("Muralidharan.Parthasarathy@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Catalog & Mapping Team -  Late Login Summary (" + indianTime.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 18) //Ops Support
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Rajesh.Subramanyam@tangoe.com"));
                    message1.To.Add(new MailAddress("Shabeenaz1@tangoe.com"));
                    message1.To.Add(new MailAddress("Muralidharan.Parthasarathy@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.Subject = "Ops Support Team -  Late Login Summary (" + indianTime.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 16) //IN_Corelogic
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Sriram.Krishnan@tangoe.com"));
                    message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));

                    message1.Subject = "IN_Corelogic Team -  Late Login Summary (" + indianTime.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }

                if (teamid == 120) //IN_Admin
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Sriram.Krishnan@tangoe.com"));
                    message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));

                    message1.Subject = "IN_Admin Team -  Late Login Summary (" + indianTime.ToString("MM-dd-yyyy") + ")-" + fromTime + " to " + toTime;
                }



                message1.Body = sb.ToString();
                message1.IsBodyHtml = true;
                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                // smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }

        }

        //Resource Utilization               

        private DataSet GetEmployees()
        {
            if (ds.Tables.Contains("users"))
            {
                ds.Tables.Remove(ds.Tables["users"]);
            }
            da = new SqlDataAdapter("SELECT UL_User_Name, UL_Team_Id FROM RTM_User_List where UL_User_Status =1", con);
            da.Fill(ds, "users");
            return ds;
        }

        private DataSet GetLoginTime(string user)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("Login"))
            {
                ds.Tables.Remove(ds.Tables["Login"]);
            }
            da = new SqlDataAdapter("SELECT LA_Start_Date_Time From RTM_Log_Actions where LA_User_Name = '" + user + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_Start_Date_Time))) BETWEEN '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToShortDateString() + "' and '" + indianTime.AddDays(-(int)indianTime.DayOfWeek).ToShortDateString() + "' and LA_Log_Action = 'First Activity' order by LA_Start_Date_Time", con);
            da.Fill(ds, "Login");
            return ds;
        }

        private DataSet GetLogoutTime(string user, DateTime date1)
        {
            if (ds.Tables.Contains("Logout"))
            {
                ds.Tables.Remove(ds.Tables["Logout"]);
            }
            da = new SqlDataAdapter("SELECT TOP 1 LA_Start_Date_Time From RTM_Log_Actions where LA_User_Name = '" + user + "' and LA_Start_Date_Time >= '" + date1 + "' and LA_Log_Action = 'Last Activity' order by LA_Start_Date_Time", con);
            da.Fill(ds, "Logout");
            return ds;
        }

        private DataSet GetTaskHours(string user, DateTime start, DateTime end)
        {
            if (ds.Tables.Contains("Task"))
            {
                ds.Tables.Remove(ds.Tables["Task"]);
            }
            da = new SqlDataAdapter("SELECT sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600 as hour, (sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/60)%60 as minute,(sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/60)%60)))%60 as seconds from RTM_Records , RTM_SubTask_List where R_SubTask = STL_ID and R_User_Name='" + user + "' and R_Start_Date_Time BETWEEN '" + start + "' and '" + end + "' and R_Duration != 'HH:MM:SS' and STL_SubTask != 'NON-TASK'", con);
            da.Fill(ds, "Task");
            return ds;
        }

        private DataSet GetLogHours(string user, DateTime start, DateTime end)
        {
            if (ds.Tables.Contains("Log"))
            {
                ds.Tables.Remove(ds.Tables["Log"]);
            }
            da = new SqlDataAdapter("SELECT sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/3600 as hour, (sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/60)%60 as minute,(sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/60)%60)))%60 as seconds from RTM_Log_Actions where LA_User_Name='" + user + "' and LA_Start_Date_Time BETWEEN '" + start + "' and '" + end + "' and LA_Duration != 'HH:MM:SS' and LA_Reason != 'Break' and LA_Reason != 'Non-Task' and  LA_Reason !='Idle Time'", con);
            da.Fill(ds, "Log");
            return ds;
        }

        private DataSet GetLeaves(string user)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("Leaves"))
            {
                ds.Tables.Remove(ds.Tables["Leaves"]);
            }
            da = new SqlDataAdapter("SELECT CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_Start_Date_Time))) From RTM_Log_Actions where LA_User_Name = '" + user + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_Start_Date_Time))) BETWEEN '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToShortDateString() + "' and '" + indianTime.AddDays(-(int)indianTime.DayOfWeek).ToShortDateString() + "' and LA_Log_Action = 'First Activity' GROUP BY CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_Start_Date_Time)))", con);
            da.Fill(ds, "Leaves");
            return ds;
        }

        private DataSet GetNonTaskRecords(string user, DateTime start, DateTime end)
        {
            if (ds.Tables.Contains("NonTaskRecords"))
            {
                ds.Tables.Remove(ds.Tables["NonTaskRecords"]);
            }
            da = new SqlDataAdapter("SELECT sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600 as hour, (sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/60)%60 as minute,(sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/60)%60)))%60 as seconds from RTM_Records, RTM_SubTask_List where R_SubTask = STL_ID and R_User_Name='" + user + "' and R_Start_Date_Time BETWEEN '" + start + "' and '" + end + "' and R_Duration != 'HH:MM:SS' and STL_SubTask = 'Non Task'", con);
            da.Fill(ds, "NonTaskRecords");
            return ds;
        }

        private DataSet GetNonTaskLogs(string user, DateTime start, DateTime end)
        {
            if (ds.Tables.Contains("NonTaskLog"))
            {
                ds.Tables.Remove(ds.Tables["NonTaskLog"]);
            }
            da = new SqlDataAdapter("SELECT sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/3600 as hour, (sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/60)%60 as minute,(sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/60)%60)))%60 as seconds from RTM_Log_Actions where LA_User_Name='" + user + "' and LA_Start_Date_Time BETWEEN '" + start + "' and '" + end + "' and LA_Duration != 'HH:MM:SS' and LA_Reason = 'Others'", con);
            da.Fill(ds, "NonTaskLog");
            return ds;
        }

        private DataSet CheckRecord(string user)
        {
            if (ds.Tables.Contains("check"))
            {
                ds.Tables.Remove(ds.Tables["check"]);
            }
            da = new SqlDataAdapter("select RU_UserName from RTM_ResourceUtil where RU_UserName='" + user + "'", con);
            da.Fill(ds, "check");
            return ds;
        }

        private void truncateTable()
        {
            cmd = new SqlCommand("Truncate table RTM_ResourceUtil", con);
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }

        private void loadRUDetails()
        {
            try
            {
                //DataRow dr;
                DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
                double totalTaskHours = 0;
                double totalLogHours = 0;
                double totalWorkHours;
                double RUPercent = 0;
                int leaveHours = 0;
                double nonTaskRecord = 0;
                double nonTaskLog = 0;
                double totalNonTaskHours = 0;

                //BuildRUTable();
                truncateTable();
                ds = GetEmployees();

                if (ds.Tables["users"].Rows.Count > 0)
                {
                    foreach (DataRow dr1 in ds.Tables["users"].Rows)
                    {
                        if (ds.Tables.Contains("Login"))
                        {
                            ds.Tables.Remove(ds.Tables["Login"]);
                        }
                        if (ds.Tables.Contains("check"))
                        {
                            ds.Tables.Remove(ds.Tables["check"]);
                        }
                        if (ds.Tables.Contains("Logout"))
                        {
                            ds.Tables.Remove(ds.Tables["Logout"]);
                        }
                        if (ds.Tables.Contains("Task"))
                        {
                            ds.Tables.Remove(ds.Tables["Task"]);
                        }
                        if (ds.Tables.Contains("Log"))
                        {
                            ds.Tables.Remove(ds.Tables["Log"]);
                        }
                        if (ds.Tables.Contains("Leaves"))
                        {
                            ds.Tables.Remove(ds.Tables["Leaves"]);
                        }
                        if (ds.Tables.Contains("NonTaskRecords"))
                        {
                            ds.Tables.Remove(ds.Tables["NonTaskRecords"]);
                        }
                        if (ds.Tables.Contains("NonTaskLog"))
                        {
                            ds.Tables.Remove(ds.Tables["NonTaskLog"]);
                        }

                        totalTaskHours = 0;
                        totalLogHours = 0;
                        totalWorkHours = 0;
                        RUPercent = 0;
                        //leaveHours = 0;
                        nonTaskRecord = 0;
                        nonTaskLog = 0;
                        totalNonTaskHours = 0;
                        string username = dr1["UL_User_Name"].ToString();
                        int teamid = Convert.ToInt32(dr1["UL_Team_Id"]);
                        DateTime logoutTime;
                        //if (username == "Jijitha Ganesh")
                        //{
                        //    username = "Jijitha Ganesh";
                        //}
                        ds = GetLoginTime(username);

                        if (ds.Tables["Login"].Rows.Count > 0)
                        {
                            DateTime loginTime = Convert.ToDateTime(ds.Tables["Login"].Rows[0]["LA_Start_Date_Time"]);

                            DataRow lastRow = ds.Tables["Login"].Rows[ds.Tables["Login"].Rows.Count - 1];


                            ds = GetLogoutTime(username, Convert.ToDateTime(lastRow["LA_Start_Date_Time"]));

                            if (ds.Tables["Logout"].Rows.Count > 0)
                            {
                                logoutTime = Convert.ToDateTime(ds.Tables["Logout"].Rows[0]["LA_Start_Date_Time"]);

                                ds = GetTaskHours(username, loginTime, logoutTime);

                                if (ds.Tables["Task"].Rows.Count > 0 && ds.Tables["Task"].Rows[0]["hour"].ToString().Length > 0)
                                {
                                    totalTaskHours = Convert.ToDouble(ds.Tables["Task"].Rows[0]["hour"]) + (Convert.ToDouble(ds.Tables["Task"].Rows[0]["minute"]) / 60) + (Convert.ToDouble(ds.Tables["Task"].Rows[0]["seconds"]) / 3600);
                                }


                                ds = GetLogHours(username, loginTime, logoutTime);

                                if (ds.Tables["Log"].Rows.Count > 0 && ds.Tables["Log"].Rows[0]["hour"].ToString().Length > 0)
                                {
                                    totalLogHours = Convert.ToDouble(ds.Tables["Log"].Rows[0]["hour"]) + (Convert.ToDouble(ds.Tables["Log"].Rows[0]["minute"]) / 60) + (Convert.ToDouble(ds.Tables["Log"].Rows[0]["seconds"]) / 3600);
                                }

                                totalWorkHours = totalTaskHours + totalLogHours;
                                totalWorkHours = Math.Round(totalWorkHours, 2, MidpointRounding.AwayFromZero);

                                ds = GetLeaves(username);
                                if (ds.Tables["Leaves"].Rows.Count <= 5)
                                {
                                    leaveHours = (5 - Convert.ToInt32(ds.Tables["Leaves"].Rows.Count)) * 8;
                                }

                                if (totalWorkHours != 0)
                                {
                                    RUPercent = (totalWorkHours / (40 - leaveHours)) * 100;
                                    RUPercent = Math.Round(RUPercent, 2, MidpointRounding.AwayFromZero);
                                }

                                ds = GetNonTaskRecords(username, loginTime, logoutTime);

                                if (ds.Tables["NonTaskRecords"].Rows.Count > 0 && ds.Tables["NonTaskRecords"].Rows[0]["hour"].ToString().Length > 0)
                                {
                                    nonTaskRecord = Convert.ToDouble(ds.Tables["NonTaskRecords"].Rows[0]["hour"] + "." + ds.Tables["NonTaskRecords"].Rows[0]["minute"]);
                                }

                                ds = GetNonTaskLogs(username, loginTime, logoutTime);

                                if (ds.Tables["NonTaskLog"].Rows.Count > 0 && ds.Tables["NonTaskLog"].Rows[0]["hour"].ToString().Length > 0)
                                {
                                    nonTaskLog = Convert.ToDouble(ds.Tables["NonTaskLog"].Rows[0]["hour"] + "." + ds.Tables["NonTaskLog"].Rows[0]["minute"]);
                                }

                                totalNonTaskHours = nonTaskRecord + nonTaskLog;

                                //dr = dtResult.NewRow();

                                //dr["Employee Name"] = username;
                                //dr["Team Name"] = TeamName;
                                //dr["Total Hours"] = totalWorkHours;
                                //dr["Absent Hours"] = leaveHours;
                                //dr["RU%"] = RUPercent;
                                //dr["NonTask"] = totalNonTaskHours;

                                //dtResult.Rows.Add(dr);

                                //ds = CheckRecord(username);
                                //if (ds.Tables["check"].Rows.Count > 0)
                                //{
                                //    cmd = new SqlCommand("update RTM_ResourceUtil set RU_TeamId=" + teamid + ", RU_WorkHours='" + totalWorkHours + "', RU_Percent='" + RUPercent + "', RU_NonTask='" + totalNonTaskHours + "', RU_StartDate= '" + DateTime.Now.AddDays(-(int)DateTime.Now.DayOfWeek - 6).ToShortDateString() + "', RU_EndDate= '" + DateTime.Now.AddDays(-(int)DateTime.Now.DayOfWeek).ToShortDateString() + "' where RU_UserName = '"+ username +"'", con);
                                //    con.Open();
                                //    cmd.ExecuteNonQuery();
                                //    con.Close();
                                //}
                                //else
                                //{
                                cmd = new SqlCommand("insert into RTM_ResourceUtil (RU_UserName, RU_TeamId, RU_WorkHours, RU_Percent, RU_NonTask, RU_StartDate, RU_EndDate, RU_Leaves) values ('" + username + "', " + teamid + ", '" + totalWorkHours + "', '" + RUPercent + "', '" + totalNonTaskHours + "', '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToShortDateString() + "', '" + indianTime.AddDays(-(int)indianTime.DayOfWeek).ToShortDateString() + "', '" + leaveHours + "')", con);
                                con.Open();
                                cmd.ExecuteNonQuery();
                                con.Close();
                                // }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                MailMessage message1 = new MailMessage();
                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress("RTM-Mailer@tangoe.com");
                //message1.To.Add(new MailAddress("Lokesha.B@tangoe.com"));
                //foreach (string item in ToAddrSuccess)
                //{
                message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                //}
                message1.Subject = "Resource Utilization script failed to execute successfully. Chart is unavailable on dashboard.";
                message1.Body = "This is a system generated mail. Please donot reply";
                message1.IsBodyHtml = false;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                //smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;
                //smtp.UseDefaultCredentials = true;
                //smtp.Credentials = new NetworkCredential("BLR-RTM-Server@tangoe.com", "");
                smtp.Send(message1);

            }
        }

        //Delay hours

        private void BuildDelayTable()
        {
            dtResult = new DataTable();
            DataColumn dc;

            dtResult.Columns.Add("Sl.No");

            dc = new DataColumn("Date");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Nature");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Team");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Employee Name");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Employee Id");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Employee's Cell no");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Scheduled Log in time");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Actual Log in time (due to delay)");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Scheduled Log out");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Actual Log out");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Delayed Login");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Delayed Log out");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Comments");
            dtResult.Columns.Add(dc);
        }

        private void BuildDayWiseDelayTable()
        {
            dtResult = new DataTable();
            DataColumn dc;

            dtResult.Columns.Add("Sl.No.");
            dc = new DataColumn("Date");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Day");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("No Of Employees");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Delayed Login Hours");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Average Delay Login Hours Per Employee");
            dtResult.Columns.Add(dc);
        }

        private DataTable GetDelayDetails()
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            dt = new DataTable();
            da = new SqlDataAdapter("select ROW_NUMBER() OVER (Order by T_TeamName) as [Sl.No], CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, D_Date))) as D_Date, D_UserName, UL_Employee_Id, UL_SCH_Login, UL_SCH_Logout, T_TeamName, D_SLogin, D_SLogout  from RTM_DelayedLogInOff left join RTM_User_List on D_UserName = UL_User_Name left join RTM_Team_List on D_Team_Id = T_ID where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, D_Date))) BETWEEN '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToShortDateString() + "' and '" + indianTime.AddDays(-(int)indianTime.DayOfWeek).ToShortDateString() + "' and D_Reason='Company Cab' Group by CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, D_Date))), D_UserName, UL_Employee_Id, UL_SCH_Login, UL_SCH_Logout, T_TeamName, D_SLogin, D_SLogout order by T_TeamName", con);
            da.Fill(dt);
            return dt;
        }

        private DataSet GetDelayLogin(string user, string date1)
        {
            if (ds.Tables.Contains("loginDelay"))
            {
                ds.Tables.Remove(ds.Tables["loginDelay"]);
            }
            //da = new SqlDataAdapter("select * from RTM_DelayedLogInOff where D_UserName = '" + user + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, D_Date))) = '" + date1 + "' and D_Type = 'In'", con);
            da = new SqlDataAdapter("select * from RTM_DelayedLogInOff, RTM_Log_Actions where D_UserName = LA_User_Name and D_UserName = '" + user + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, D_Date))) = '" + date1 + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_Start_Date_Time))) = '" + date1 + "' and D_Type = 'In' and LA_Log_Action='First Activity'", con);
            da.Fill(ds, "loginDelay");
            return ds;
        }

        private DataSet GetDelayLogout(string user, string date1)
        {
            if (ds.Tables.Contains("logoutDelay"))
            {
                ds.Tables.Remove(ds.Tables["logoutDelay"]);
            }
            //da = new SqlDataAdapter("select * from RTM_DelayedLogInOff where D_UserName = '" + user + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, D_Date))) = '" + date1 + "' and D_Type = 'Off'", con);
            da = new SqlDataAdapter("select * from RTM_DelayedLogInOff , RTM_Log_Actions where D_UserName = LA_User_Name and D_UserName = '" + user + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, D_Date))) = '" + date1 + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_Start_Date_Time))) = '" + date1 + "' and D_Type = 'Off' and LA_Log_Action='Last Activity'", con);
            da.Fill(ds, "logoutDelay");
            return ds;
        }

        private DataSet GetDayWiseDelay()
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("dayWiseDelay"))
            {
                ds.Tables.Remove(ds.Tables["dayWiseDelay"]);
            }

            string query = "select ROW_NUMBER() OVER (Order by Day(D_Date)) as [Sl.No], Day(D_Date), CONVERT(VARCHAR(10),D_Date,101) as [Date], datename(dw,CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, D_Date)))) as [Day], count(Distinct D_UserName) as [UserCount], CONVERT(varchar(10), sum(datediff(second,'00:00:00',REPLACE(D_Duration,'-', '')))/3600) as [DelayedHours] from RTM_DelayedLogInOff where D_Reason='Company Cab' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, D_Date))) between '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToShortDateString() + "' and '" + indianTime.AddDays(-(int)indianTime.DayOfWeek).ToShortDateString() + "' " +
                            "Group by Day(D_Date), datename(dw,CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, D_Date)))), CONVERT(VARCHAR(10),D_Date,101) " +
                            "Order BY Day(D_Date)";
            da = new SqlDataAdapter(query, con);
            da.Fill(ds, "dayWiseDelay");
            return ds;
        }

        private void BindDelayDetails()
        {
            try
            {
                DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
                DataRow dr;

                BuildDelayTable();

                dt = GetDelayDetails();

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr1 in dt.Rows)
                    {
                        dr = dtResult.NewRow();
                        dr["Sl.No"] = dr1["Sl.No"];
                        dr["Date"] = Convert.ToDateTime(dr1["D_Date"]).ToShortDateString();
                        dr["Nature"] = "Company Transport";
                        dr["Team"] = dr1["T_TeamName"].ToString();
                        dr["Employee Name"] = dr1["D_UserName"];
                        dr["Employee Id"] = dr1["UL_Employee_Id"];
                        dr["Employee's Cell no"] = "";
                        dr["Scheduled Log in time"] = dr1["D_SLogin"] == DBNull.Value ? "" : Convert.ToDateTime(dr1["D_SLogin"]).ToShortTimeString();

                        ds = GetDelayLogin(dr1["D_UserName"].ToString(), Convert.ToDateTime(dr1["D_Date"]).ToShortDateString());

                        if (ds.Tables["loginDelay"].Rows.Count > 0)
                        {
                            //dr["Actual Log in time (due to delay)"] = ds.Tables["loginDelay"].Rows[0]["D_Date"] == DBNull.Value ? "" : Convert.ToDateTime(ds.Tables["loginDelay"].Rows[0]["D_Date"]).ToShortTimeString();
                            //dr["Delayed Login"] = TimeSpan.Parse(ds.Tables["loginDelay"].Rows[0]["D_Duration"].ToString()).TotalMinutes;
                            //dr["Delayed Login"] = decimal.Round(Convert.ToDecimal(dr["Delayed Login"]), 0, MidpointRounding.AwayFromZero);

                            dr["Actual Log in time (due to delay)"] = ds.Tables["loginDelay"].Rows[0]["LA_Start_Date_Time"] == DBNull.Value ? "" : Convert.ToDateTime(ds.Tables["loginDelay"].Rows[0]["LA_Start_Date_Time"]).ToShortTimeString();
                            TimeSpan delayloginspan = Convert.ToDateTime(ds.Tables["loginDelay"].Rows[0]["LA_Start_Date_Time"]).TimeOfDay.Subtract(Convert.ToDateTime(dr1["D_SLogin"]).TimeOfDay);
                            dr["Delayed Login"] = delayloginspan.TotalMinutes;
                            dr["Delayed Login"] = decimal.Round(Convert.ToDecimal(dr["Delayed Login"]), 0, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            dr["Actual Log in time (due to delay)"] = "";
                            dr["Delayed Login"] = "";
                        }

                        dr["Scheduled Log out"] = dr1["D_SLogout"] == DBNull.Value ? "" : Convert.ToDateTime(dr1["D_SLogout"]).ToShortTimeString();

                        ds = GetDelayLogout(dr1["D_UserName"].ToString(), Convert.ToDateTime(dr1["D_Date"]).ToShortDateString());

                        if (ds.Tables["logoutDelay"].Rows.Count > 0)
                        {
                            //dr["Actual Log out"] = ds.Tables["logoutDelay"].Rows[0]["D_Date"] == DBNull.Value ? "" : Convert.ToDateTime(ds.Tables["logoutDelay"].Rows[0]["D_Date"]).ToShortTimeString();
                            //dr["Delayed Log out"] = TimeSpan.Parse(ds.Tables["logoutDelay"].Rows[0]["D_Duration"].ToString()).TotalMinutes;
                            //dr["Delayed Log out"] = decimal.Round(Convert.ToDecimal(dr["Delayed Log out"]), 0, MidpointRounding.AwayFromZero);

                            dr["Actual Log out"] = ds.Tables["logoutDelay"].Rows[0]["LA_Start_Date_Time"] == DBNull.Value ? "" : Convert.ToDateTime(ds.Tables["logoutDelay"].Rows[0]["LA_Start_Date_Time"]).ToShortTimeString();
                            TimeSpan delaylogoutspan = Convert.ToDateTime(ds.Tables["logoutDelay"].Rows[0]["LA_Start_Date_Time"]).TimeOfDay.Subtract(Convert.ToDateTime(dr1["D_SLogout"]).TimeOfDay);
                            dr["Delayed Log out"] = delaylogoutspan.TotalMinutes;
                            dr["Delayed Log out"] = decimal.Round(Convert.ToDecimal(dr["Delayed Log out"]), 0, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            dr["Actual Log out"] = "";
                            dr["Delayed Log out"] = "";
                        }

                        dr["Comments"] = "";

                        dtResult.Rows.Add(dr);
                    }

                    MailMessage message1 = new MailMessage();
                    //string filePath = "" + Directory.GetCurrentDirectory() + "\\Transport Delay – " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToString("MM-dd-yyyy") + "To" + indianTime.AddDays(-(int)indianTime.DayOfWeek).ToString("MM-dd-yyyy") + ".csv";
                    //CSVUtility.ToCSV(dtResult, filePath);
                    DirectCSV csv = new DirectCSV();
                    var data = csv.ExportToCSV(dtResult);
                    var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(data);
                    MemoryStream ms = new MemoryStream(bytes);

                    Attachment attachFile = new Attachment(ms, "Transport Delay – " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToString("MM-dd-yyyy") + "To" + indianTime.AddDays(-(int)indianTime.DayOfWeek).ToString("MM-dd-yyyy") + ".csv", "application/csv");
                    message1.Attachments.Add(attachFile);

                    int totalDelay = 0;
                    double totalAverage = 0;

                    BuildDayWiseDelayTable();

                    ds = GetDayWiseDelay();
                    if (ds.Tables["dayWiseDelay"].Rows.Count > 0)
                    {
                        foreach (DataRow drRow in ds.Tables["dayWiseDelay"].Rows)
                        {
                            dr = dtResult.NewRow();
                            dr["Sl.No."] = drRow["Sl.No"];
                            dr["Date"] = drRow["Date"];
                            dr["Day"] = drRow["Day"];
                            dr["No Of Employees"] = drRow["UserCount"];
                            dr["Delayed Login Hours"] = drRow["DelayedHours"];
                            if (drRow["DelayedHours"].ToString().Length > 0)
                            {
                                totalDelay = totalDelay + Convert.ToInt32(drRow["DelayedHours"]);
                                totalAverage = totalAverage + Math.Round((Convert.ToDouble(drRow["DelayedHours"]) / Convert.ToDouble(drRow["UserCount"])), 2, MidpointRounding.AwayFromZero);
                            }


                            dr["Average Delay Login Hours Per Employee"] = Math.Round((Convert.ToDouble(drRow["DelayedHours"]) / Convert.ToDouble(drRow["UserCount"])), 2, MidpointRounding.AwayFromZero);

                            dtResult.Rows.Add(dr);
                        }
                    }

                    //ds = GetDayWiseDelay("Monday");
                    //if (ds.Tables["dayWiseDelay"].Rows.Count > 0)
                    //{
                    //    dr = dtResult.NewRow();
                    //    dr["Date"] = indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToShortDateString();
                    //    dr["Day"] = "Monday";
                    //    dr["No Of Employees"] = ds.Tables["dayWiseDelay"].Rows[0]["UserCount"];
                    //    dr["Delayed Login Hours"] = ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"];
                    //    if (ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"].ToString().Length > 0)
                    //    {
                    //        totalDelay = totalDelay + Convert.ToInt32(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]);
                    //        totalAverage = totalAverage + Math.Round((Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]) / Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["UserCount"])), 2, MidpointRounding.AwayFromZero);
                    //    }


                    //    dr["Average Delay Login Hours Per Employee"] = Math.Round((Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]) / Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["UserCount"])), 2, MidpointRounding.AwayFromZero);

                    //    dtResult.Rows.Add(dr);
                    //}

                    //ds = GetDayWiseDelay("Tuesday");
                    //if (ds.Tables["dayWiseDelay"].Rows.Count > 0)
                    //{
                    //    dr = dtResult.NewRow();
                    //    dr["Date"] = indianTime.AddDays(-(int)indianTime.DayOfWeek - 5).ToShortDateString();
                    //    dr["Day"] = "Tuesday";
                    //    dr["No Of Employees"] = ds.Tables["dayWiseDelay"].Rows[0]["UserCount"];
                    //    dr["Delayed Login Hours"] = ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"];
                    //    if (ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"].ToString().Length > 0)
                    //    {
                    //        totalDelay = totalDelay + Convert.ToInt32(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]);
                    //        totalAverage = totalAverage + Math.Round((Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]) / Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["UserCount"])), 2, MidpointRounding.AwayFromZero);
                    //    }
                    //    //totalDelay = totalDelay + Convert.ToInt32(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]);
                    //    // totalAverage = totalAverage + Math.Round((Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]) / Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["UserCount"])), 2,MidpointRounding.AwayFromZero );
                    //    dr["Average Delay Login Hours Per Employee"] = Math.Round((Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]) / Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["UserCount"])), 2, MidpointRounding.AwayFromZero);

                    //    dtResult.Rows.Add(dr);
                    //}

                    //ds = GetDayWiseDelay("Wednesday");
                    //if (ds.Tables["dayWiseDelay"].Rows.Count > 0)
                    //{
                    //    dr = dtResult.NewRow();
                    //    dr["Date"] = indianTime.AddDays(-(int)indianTime.DayOfWeek - 4).ToShortDateString();
                    //    dr["Day"] = "Wednesday";
                    //    dr["No Of Employees"] = ds.Tables["dayWiseDelay"].Rows[0]["UserCount"];
                    //    dr["Delayed Login Hours"] = ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"];
                    //    if (ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"].ToString().Length > 0)
                    //    {
                    //        totalDelay = totalDelay + Convert.ToInt32(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]);
                    //        totalAverage = totalAverage + Math.Round((Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]) / Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["UserCount"])), 2, MidpointRounding.AwayFromZero);
                    //    }
                    //    //totalDelay = totalDelay + Convert.ToInt32(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]);
                    //    //totalAverage = totalAverage + Math.Round((Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]) / Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["UserCount"])), 2,MidpointRounding.AwayFromZero );
                    //    dr["Average Delay Login Hours Per Employee"] = Math.Round((Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]) / Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["UserCount"])), 2, MidpointRounding.AwayFromZero);

                    //    dtResult.Rows.Add(dr);
                    //}

                    //ds = GetDayWiseDelay("Thursday");
                    //if (ds.Tables["dayWiseDelay"].Rows.Count > 0)
                    //{
                    //    dr = dtResult.NewRow();
                    //    dr["Date"] = indianTime.AddDays(-(int)indianTime.DayOfWeek - 3).ToShortDateString();
                    //    dr["Day"] = "Thursday";
                    //    dr["No Of Employees"] = ds.Tables["dayWiseDelay"].Rows[0]["UserCount"];
                    //    dr["Delayed Login Hours"] = ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"];
                    //    if (ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"].ToString().Length > 0)
                    //    {
                    //        totalDelay = totalDelay + Convert.ToInt32(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]);
                    //        totalAverage = totalAverage + Math.Round((Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]) / Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["UserCount"])), 2, MidpointRounding.AwayFromZero);
                    //    }
                    //    //totalDelay = totalDelay + Convert.ToInt32(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]);
                    //    //totalAverage = totalAverage + Math.Round((Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]) / Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["UserCount"])), 2,MidpointRounding.AwayFromZero );
                    //    dr["Average Delay Login Hours Per Employee"] = Math.Round((Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]) / Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["UserCount"])), 2, MidpointRounding.AwayFromZero);

                    //    dtResult.Rows.Add(dr);
                    //}

                    //ds = GetDayWiseDelay("Friday");
                    //if (ds.Tables["dayWiseDelay"].Rows.Count > 0)
                    //{
                    //    dr = dtResult.NewRow();
                    //    dr["Date"] = indianTime.AddDays(-(int)indianTime.DayOfWeek - 2).ToShortDateString();
                    //    dr["Day"] = "Friday";

                    //    dr["No Of Employees"] = ds.Tables["dayWiseDelay"].Rows[0]["UserCount"];
                    //    dr["Delayed Login Hours"] = ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"];
                    //    if (ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"].ToString().Length > 0)
                    //    {
                    //        totalDelay = totalDelay + Convert.ToInt32(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]);
                    //        totalAverage = totalAverage + Math.Round((Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]) / Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["UserCount"])), 2, MidpointRounding.AwayFromZero);
                    //    }
                    //    // totalDelay = totalDelay + Convert.ToInt32(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]);
                    //    // totalAverage = totalAverage + Math.Round((Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]) / Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["UserCount"])), 2,MidpointRounding.AwayFromZero );
                    //    dr["Average Delay Login Hours Per Employee"] = Math.Round((Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["DelayedHours"]) / Convert.ToDouble(ds.Tables["dayWiseDelay"].Rows[0]["UserCount"])), 2, MidpointRounding.AwayFromZero);

                    //    dtResult.Rows.Add(dr);
                    //}

                    dr = dtResult.NewRow();

                    dr["Date"] = "";
                    dr["Day"] = "Total";
                    dr["No Of Employees"] = "";
                    dr["Delayed Login Hours"] = totalDelay;
                    dr["Average Delay Login Hours Per Employee"] = totalAverage / 5;

                    dtResult.Rows.Add(dr);

                    getDelayHTML(dtResult);

                    StringBuilder sb = new StringBuilder();

                    sb.AppendLine("Hi All,");
                    sb.AppendLine("");
                    sb.AppendLine("Please find the attached report for Cab delays for the last week.");
                    sb.AppendLine("");
                    sb.AppendLine(myBuilder.ToString());   //here I want the data to       display in table format
                    sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                    sb.AppendLine("");


                    SmtpClient smtp = new SmtpClient();

                    message1.From = new MailAddress(FromAddress);
                    //message1.To.Add(new MailAddress("Lokesha.B@tangoe.com"));
                    foreach (string item in ToAddrSuccess)
                    {
                        message1.To.Add(new MailAddress(item));
                    }
                    message1.Subject = "Transport Issues – Delayed Login and Logoff – ( " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToString("MM-dd-yyyy") + " – " + indianTime.AddDays(-(int)indianTime.DayOfWeek).ToString("MM-dd-yyyy") + ")";
                    //message1.Body = "Hi All, " + Environment.NewLine + "Please find the attached report for Cab delays for the last week." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.";
                    message1.Body = sb.ToString();

                    //System.Net.Mail.Attachment attachment;
                    //attachment = new System.Net.Mail.Attachment(filePath);
                    //message1.Attachments.Add(attachment);
                    message1.IsBodyHtml = true;

                    smtp.Port = 25;
                    smtp.Host = "10.0.5.104";
                    //smtp.Host = "outlook-south.tangoe.com";
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    smtp.EnableSsl = false;
                    //smtp.UseDefaultCredentials = true;
                    //smtp.Credentials = new NetworkCredential("BLR-RTM-Server@tangoe.com", "");
                    smtp.Send(message1);
                }
            }
            catch (Exception)
            {

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
                if (myRow.ItemArray.Contains("Total"))
                {
                    myBuilder.Append("<tr align='left' valign='top'>");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        myBuilder.Append("<td align='left' valign='top' bgcolor='#FFFF00'>");
                        myBuilder.Append("<B />" + myRow[myColumn.ColumnName].ToString());
                        myBuilder.Append("</td>");
                    }
                    myBuilder.Append("</tr>");
                }
                else
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

            }
            myBuilder.Append("</table>");

            return myBuilder.ToString();
        }

        private string getLateLoginHTML(DataTable dt)
        {
            myBuilder = new StringBuilder();

            myBuilder.Append("<table border='1' cellpadding='5' cellspacing='0' ");
            myBuilder.Append("style='border: solid 1px Silver; font-size: x-small;'>");

            myBuilder.Append("<tr align='left' valign='top'>");
            foreach (DataColumn myColumn in dt.Columns)
            {
                myBuilder.Append("<td align='left' valign='top' bgcolor='#008000'>");
                myBuilder.Append("<B />" + myColumn.ColumnName);
                myBuilder.Append("</td>");
            }
            myBuilder.Append("</tr>");

            foreach (DataRow myRow in dt.Rows)
            {
                if (myRow.ItemArray.Contains("Total"))
                {
                    myBuilder.Append("<tr align='left' valign='top'>");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        myBuilder.Append("<td align='left' valign='top' bgcolor='#008000'>");
                        myBuilder.Append("<B />" + myRow[myColumn.ColumnName].ToString());
                        myBuilder.Append("</td>");
                    }
                    myBuilder.Append("</tr>");
                }
                else
                {
                    //string mins1 = myRow.ItemArray[5].ToString();
                    double mins = TimeSpan.Parse(myRow.ItemArray[5].ToString()).TotalMinutes;
                    int i = 0;
                    if (mins >= 30)
                    {
                        myBuilder.Append("<tr align='left' valign='top'>");
                        foreach (DataColumn myColumn in dt.Columns)
                        {
                            if (i == 0)
                            {
                                myBuilder.Append("<td align='left' valign='top' bgcolor='#ff0000'>");
                                myBuilder.Append(myRow[myColumn.ColumnName].ToString());
                                myBuilder.Append("</td>");
                                i = i + 1;
                            }
                            else
                            {
                                myBuilder.Append("<td align='left' valign='top'>");
                                myBuilder.Append(myRow[myColumn.ColumnName].ToString());
                                myBuilder.Append("</td>");
                            }
                        }
                        myBuilder.Append("</tr>");
                    }
                    else if (mins >= 15 && mins < 30)
                    {
                        myBuilder.Append("<tr align='left' valign='top'>");
                        foreach (DataColumn myColumn in dt.Columns)
                        {
                            if (i == 0)
                            {
                                myBuilder.Append("<td align='left' valign='top' bgcolor='#ffff00'>");
                                myBuilder.Append(myRow[myColumn.ColumnName].ToString());
                                myBuilder.Append("</td>");
                                i = i + 1;
                            }
                            else
                            {
                                myBuilder.Append("<td align='left' valign='top'>");
                                myBuilder.Append(myRow[myColumn.ColumnName].ToString());
                                myBuilder.Append("</td>");
                            }
                        }
                        myBuilder.Append("</tr>");
                    }
                    else
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

                }

            }
            myBuilder.Append("</table>");

            return myBuilder.ToString();
        }
        //OnBoard Estimate report
        private void DisplayEstRecords(int TID)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (indianTime.DayOfWeek == DayOfWeek.Saturday)
            {
                return;
            }
            if (indianTime.DayOfWeek == DayOfWeek.Sunday)
            {
                return;
            }
            dt = new DataTable();
            da = new SqlDataAdapter("select EST_UserName as UserName, CL_ClientName as Client, TL_Task as Task, STL_SubTask as SubTask, EST_Duration as Duration, CONVERT(VARCHAR(10),EST_Date,101) as EST_Date from RTM_Estimation left join rtm_client_list on EST_ClientId = CL_ID left join rtm_task_list on EST_TaskId = TL_ID " +
             "left join rtm_subtask_list on EST_SubTaskId = STL_ID where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, EST_Date))) = '" + indianTime.ToShortDateString() + "' and EST_TeamId='" + TID + "' Order By EST_UserName", con);

            da.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                string filePath = "";
                if (TID == 10)
                {
                    filePath = "Resource Utilization Estimate – Onboarding – " + indianTime.ToString("MM-dd-yyyy") + ".csv";
                }
                else if (TID == 13)
                {
                    filePath = "Resource Utilization Estimate – Implementation – " + indianTime.ToString("MM-dd-yyyy") + ".csv";
                }
                else if (TID == 31)
                {
                    filePath = "Resource Utilization Estimate – VTM Testing – " + indianTime.ToString("MM-dd-yyyy") + ".csv";
                }
                else if (TID == 9)
                {
                    filePath = "Resource Utilization Estimate – Quality Check – " + indianTime.ToString("MM-dd-yyyy") + ".csv";
                }
                else if (TID == 11)
                {
                    filePath = "Resource Utilization Estimate – Audit & Optimize – " + indianTime.ToString("MM-dd-yyyy") + ".csv";
                }
                else if (TID == 18)
                {
                    filePath = "Resource Utilization Estimate – Ops Support – " + indianTime.ToString("MM-dd-yyyy") + ".csv";
                }
                else if (TID == 22)
                {
                    filePath = "Resource Utilization Estimate – Client Services 1 – " + indianTime.ToString("MM-dd-yyyy") + ".csv";
                }
                else if (TID == 23)
                {
                    filePath = "Resource Utilization Estimate – Client Services 2 – " + indianTime.ToString("MM-dd-yyyy") + ".csv";
                }
                else if (TID == 29)
                {
                    filePath = "Resource Utilization Estimate – Invoices – " + indianTime.ToString("MM-dd-yyyy") + ".csv";
                }

                //string filePath = "" + Directory.GetCurrentDirectory() + "\\Resource Utilization Estimate – Onboarding – " + DateTime.Now.ToShortDateString() + ".csv";

                //CSVUtility.ToCSV(dt, filePath);
                MailMessage message1 = new MailMessage();
                DirectCSV csv = new DirectCSV();
                var data = csv.ExportToCSV(dt);
                var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(data);
                MemoryStream ms = new MemoryStream(bytes);

                Attachment attachFile = new Attachment(ms, filePath, "application/csv");
                message1.Attachments.Add(attachFile);

                BuildEstimateTable();


                SmtpClient smtp = new SmtpClient();
                ds = GetTotalEstimateExpected(TID);
                ds = GetActualEstimate(TID);
                DataRow dr4;
                if (ds.Tables["actual"].Rows.Count > 0)
                {
                    foreach (DataRow dr2 in ds.Tables["actual"].Rows)
                    {
                        dr4 = dt.NewRow();

                        dr4["User Name"] = dr2["User Name"];
                        dr4["Estimated Total Duration"] = dr2["Estimated Total Duration"];

                        ds = GetDelayLoginDuration(dr2["User Name"].ToString());
                        if (ds.Tables["cab"].Rows.Count > 0)
                        {
                            dr4["Cab Delay"] = ds.Tables["cab"].Rows[0]["D_Duration"].ToString();
                        }

                        dt.Rows.Add(dr4);

                    }
                }
                message1.From = new MailAddress(FromAddress);

                getHTML(ds, dt);

                StringBuilder sb = new StringBuilder();

                sb.AppendLine("Hi All,");
                sb.AppendLine("");
                sb.AppendLine("Please find the attached report for today's Resource Utilization Estimate.");
                sb.AppendLine("");
                sb.AppendLine(myBuilder.ToString());   //here I want the data to       display in table format
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                sb.AppendLine("");

                if (TID == 10)
                {
                    foreach (string item in OnBoardToAddress)
                    {
                        message1.To.Add(new MailAddress(item));
                    }
                    message1.Subject = "Onboarding Team Resource Utilization Estimate Report – ( " + indianTime.ToShortDateString() + ")";



                    message1.Body = sb.ToString(); //"Hi All, " + Environment.NewLine + "Please find the attached report for today's Resource Utilization Estimate." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.";
                }
                else if (TID == 13)
                {
                    message1.To.Add(new MailAddress("Balaji.Nagabhushanam@tangoe.com"));
                    message1.To.Add(new MailAddress("Nagabharana.Sathyanarayana@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Daily Estimate Report - Implementation Team Resource Utilization – ( " + indianTime.ToShortDateString() + ")";
                    message1.Body = sb.ToString();  //"Hi All, " + Environment.NewLine + "Please find the attached report for today's Resource Utilization Estimate." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.";

                }
                else if (TID == 31)
                {
                    message1.To.Add(new MailAddress("Sleema.Joseph@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Daily Estimate Report - VTM Testing Team Resource Utilization – ( " + indianTime.ToShortDateString() + ")";
                    message1.Body = sb.ToString();  //"Hi All, " + Environment.NewLine + "Please find the attached report for today's Resource Utilization Estimate." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.";

                }
                else if (TID == 9)
                {
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Johwessly.Chennaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Daily Estimate Report - Quality Check Team Resource Utilization – ( " + indianTime.ToShortDateString() + ")";
                    message1.Body = sb.ToString();
                }
                else if (TID == 11)
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Vinith.bekal@tangoe.com"));
                    message1.To.Add(new MailAddress("Balaji.Goalla@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Daily Estimate Report - Audit & Optimize Team Resource Utilization – ( " + indianTime.ToShortDateString() + ")";
                    message1.Body = sb.ToString();
                }
                else if (TID == 18)
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Daily Estimate Report - Ops Support Team Resource Utilization – ( " + indianTime.ToShortDateString() + ")";
                    message1.Body = sb.ToString();
                }
                else if (TID == 22)
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Vikas.Vyas@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Daily Estimate Report - Client Services 1 Team Resource Utilization – ( " + indianTime.ToShortDateString() + ")";
                    message1.Body = sb.ToString();
                }
                else if (TID == 23)
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Monnappa.Badumanda@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Daily Estimate Report - Client Services 2 Team Resource Utilization – ( " + indianTime.ToShortDateString() + ")";
                    message1.Body = sb.ToString();
                }
                else if (TID == 29)
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Vikas.Vyas@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Daily Estimate Report - Invoices Team Resource Utilization – ( " + indianTime.ToShortDateString() + ")";
                    message1.Body = sb.ToString();
                }

                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment(filePath);
                //message1.Attachments.Add(attachment);
                //message1.IsBodyHtml = false;
                message1.IsBodyHtml = true;
                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                //smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);

                //if (File.Exists(filePath))
                //{
                //    File.Delete(filePath);
                //}
            }
            else
            {
                MailMessage message1 = new MailMessage();
                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);
                if (TID == 10)
                {
                    foreach (string item in OnBoardToAddress)
                    {
                        message1.To.Add(new MailAddress(item));
                    }

                    message1.Subject = "Onboarding Team Resource Utilization Estimate Report – ( " + indianTime.ToShortDateString() + ")";
                }
                else if (TID == 13)
                {
                    message1.To.Add(new MailAddress("Balaji.Nagabhushanam@tangoe.com"));
                    message1.To.Add(new MailAddress("Nagabharana.Sathyanarayana@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Daily Estimate Report - Implementation Team Resource Utilization – ( " + indianTime.ToShortDateString() + ")";
                }
                else if (TID == 31)
                {
                    message1.To.Add(new MailAddress("Sleema.Joseph@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Daily Estimate Report - VTM Testing Team Resource Utilization – ( " + indianTime.ToShortDateString() + ")";
                }
                else if (TID == 9)
                {
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Johwessly.Chennaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Daily Estimate Report - Quality Check Team Resource Utilization – ( " + indianTime.ToShortDateString() + ")";
                }
                else if (TID == 11)
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Vinith.bekal@tangoe.com"));
                    message1.To.Add(new MailAddress("Balaji.Goalla@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Daily Estimate Report - Audit & Optimize Team Resource Utilization – ( " + indianTime.ToShortDateString() + ")";
                }
                else if (TID == 18)
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Daily Estimate Report - Ops Support Team Resource Utilization – ( " + indianTime.ToShortDateString() + ")";
                }
                else if (TID == 22)
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Vikas.Vyas@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Daily Estimate Report - Client Services 1 Team Resource Utilization – ( " + indianTime.ToShortDateString() + ")";

                }
                else if (TID == 23)
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Monnappa.Badumanda@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Daily Estimate Report - Client Services 2 Team Resource Utilization – ( " + indianTime.ToShortDateString() + ")";

                }
                else if (TID == 29)
                {
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("Vikas.Vyas@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Daily Estimate Report - Invoices Team Resource Utilization – ( " + indianTime.ToShortDateString() + ")";

                }

                //message1.Subject = "Team Resource Utilization Estimate Report – ( " + DateTime.Now.ToShortDateString() + ")";
                message1.Body = "Hi All, " + Environment.NewLine + "No estimate records found for the day." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.";

                message1.IsBodyHtml = false;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                //smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }


        }

        private void BuildEstimateTable()
        {
            dt = new DataTable();
            DataColumn dc;

            dc = new DataColumn("User Name");
            dt.Columns.Add(dc);

            dc = new DataColumn("Estimated Total Duration");
            dt.Columns.Add(dc);

            dc = new DataColumn("Cab Delay");
            dt.Columns.Add(dc);
        }

        private DataSet GetTotalEstimateExpected(int TID)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("expected"))
            {
                ds.Tables.Remove(ds.Tables["expected"]);
            }
            da = new SqlDataAdapter("select COUNT(Distinct EST_UserName) as [User Count], " +
                                "CONVERT(varchar(10), sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/3600) +':'+ " +
                                "CONVERT(varchar(10),(sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/60)%60) +':'+ " +
                                "CONVERT(varchar(10),(sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/60)%60)))%60) as [Estimated Total Duration], " +
                                "CONVERT(varchar(10),COUNT(Distinct EST_UserName) * 8) +':00:00' as [Expected Total Duration] " +
                                "from RTM_Estimation where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, EST_Date)))='" + indianTime.ToShortDateString() + "' and EST_TeamId='" + TID + "'", con);
            da.Fill(ds, "expected");
            return ds;
        }

        private DataSet GetActualEstimate(int TID)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("actual"))
            {
                ds.Tables.Remove(ds.Tables["actual"]);
            }
            da = new SqlDataAdapter("select EST_UserName as [User Name], " +
                                    "CONVERT(varchar(10), sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/3600) +':'+ " +
                                    "CONVERT(varchar(10),(sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/60)%60) +':'+ " +
                                    "CONVERT(varchar(10),(sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/60)%60)))%60) as [Estimated Total Duration] " +
                                    "from RTM_Estimation where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, EST_Date)))='" + indianTime.ToShortDateString() + "' and EST_TeamId='" + TID + "' Group By EST_UserName HAVING sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/3600 < 8", con);
            da.Fill(ds, "actual");
            return ds;
        }

        private DataSet GetDelayLoginDuration(string user)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("cab"))
            {
                ds.Tables.Remove(ds.Tables["cab"]);
            }

            da = new SqlDataAdapter("select D_Duration from dbo.RTM_DelayedLogInOff where D_UserName= '" + user + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, D_Date)))='" + indianTime.ToShortDateString() + "' and D_Reason='Company Cab'", con);
            da.Fill(ds, "cab");
            return ds;
        }

        private string getHTML(DataSet ds1, DataTable dt)
        {
            myBuilder = new StringBuilder();

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

        //OnBoard Estimate comparison Report

        private void BuidTable()
        {
            dtResult = new DataTable();
            DataColumn dc;

            dc = new DataColumn("Employee Name");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Date");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Client");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Task");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Subtask");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("User Estimate");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("RTM Time");
            dtResult.Columns.Add(dc);
        }

        private DataSet GetRTMRecords(int TID)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("Records"))
            {
                ds.Tables.Remove(ds.Tables["Records"]);
            }

            da = new SqlDataAdapter("select R_User_Name, R_Client, CL_ClientName, R_Task, TL_Task, R_SubTask, STL_SubTask, sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600 as hour, (sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/60)%60 as minute,(sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/60)%60)))%60 as seconds " +
                         " from RTM_Records left join rtm_client_list on R_Client = CL_ID left join rtm_task_list on R_Task = TL_ID " +
                        " left join rtm_subtask_list on R_SubTask = STL_ID left join RTM_User_List on R_User_Name = UL_User_Name " +
                         " where R_TeamId = " + TID + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_Start_Date_Time))) = '" + indianTime.AddDays(-1).ToShortDateString() + "' and UL_User_Status=1 " +
                          " Group by R_User_Name, R_Client, CL_ClientName, R_Task, TL_Task, R_SubTask, STL_SubTask", con);

            da.Fill(ds, "Records");
            return ds;
        }

        private DataTable GetEstimateRecords(string user, int client, int task, int subtask)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            dt = new DataTable();
            da = new SqlDataAdapter("select  sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/3600 as hour, (sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/60)%60 as minute,(sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/60)%60)))%60 as seconds " +
                " from RTM_Estimation left join rtm_client_list on EST_ClientId = CL_ID left join rtm_task_list on EST_TaskId = TL_ID" +
                " left join rtm_subtask_list on EST_SubTaskId = STL_ID where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, EST_Date))) = '" + indianTime.AddDays(-1).ToShortDateString() + "' and  EST_ClientId ='" + client + "' and EST_TaskId = '" + task + "' and EST_SubTaskId = '" + subtask + "' and EST_UserName ='" + user + "' ", con);
            da.Fill(dt);
            return dt;
        }

        private DataTable GetLogs(int TID)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            dt = new DataTable();

            da = new SqlDataAdapter("select LA_User_Name, LA_Reason, sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/3600 as hour, (sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/60)%60 as minute,(sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/60)%60)))%60 as seconds " +
            "from RTM_Log_Actions, RTM_User_List where LA_User_Name = UL_User_Name and LA_TeamId = " + TID + " and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_Start_Date_Time))) ='" + indianTime.AddDays(-1).ToShortDateString() + "' and (LA_Reason = 'Meetings' or LA_Reason='Meeting'  or LA_Reason = 'Conference-Call' ) and UL_User_Status=1  group by LA_User_Name, LA_Reason", con);

            da.Fill(dt);
            return dt;
        }

        private DataSet GetEstNotInRec(int TID)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("onlyest"))
            {
                ds.Tables.Remove(ds.Tables["onlyest"]);
            }
            da = new SqlDataAdapter("select EST_UserName, EST_ClientId,CL_ClientName, EST_TaskId,TL_Task, EST_SubTaskId, STL_SubTask, EST_Duration from RTM_Estimation left join rtm_client_list on EST_ClientId = CL_ID left join rtm_task_list on EST_TaskId = TL_ID left join rtm_subtask_list on EST_SubTaskId = STL_ID where NOT EXISTS (select * from RTM_Records where EST_ClientId = R_Client and EST_TaskId= R_Task and EST_SubTaskId=R_SubTask  and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_Start_Date_Time))) ='" + indianTime.AddDays(-1).ToShortDateString() + "' ) and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, EST_Date))) ='" + indianTime.AddDays(-1).ToShortDateString() + "' and EST_TeamId = '" + TID + "'", con);
            da.Fill(ds, "onlyest");
            return ds;
        }

        private void BuildESTCMPBodyTable()
        {
            dtResult = new DataTable();
            DataColumn dc;

            dc = new DataColumn("S.No");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Employee Name");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Date");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("User Estimate");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("RTM Task Time");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Meeting");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Conference Call");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Peer Support");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Less than 8");
            dtResult.Columns.Add(dc);
        }

        private DataSet GetEstimationsForMailBody(int TID)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("estimation"))
            {
                ds.Tables.Remove(ds.Tables["estimation"]);
            }

            da = new SqlDataAdapter("select EST_UserName as [Employee Name], CONVERT(VARCHAR(10), EST_Date, 101) as Date, " +
                    "CONVERT(Varchar(2),ISNULL(sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/3600,00)) +':'+CONVERT(Varchar(2),ISNULL((sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/60)%60,00))+':'+CONVERT(Varchar(2),ISNULL((sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(EST_Duration,'-', '')))/60)%60)))%60,00)) as [User Estimate] " +
                    "from RTM_Estimation where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, EST_Date))) ='" + indianTime.AddDays(-1).ToShortDateString() + "' and EST_TeamId = '" + TID + "'  GROUP BY EST_UserName,CONVERT(VARCHAR(10), EST_Date, 101)", con);
            da.Fill(ds, "estimation");
            return ds;
        }

        private DataSet GetActualTaskDuration(string user)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("task"))
            {
                ds.Tables.Remove(ds.Tables["task"]);
            }

            da = new SqlDataAdapter("select CONVERT(Varchar(2),ISNULL(sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600,00)) +':'+ CONVERT(Varchar(2),ISNULL((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/60)%60,00))+':'+CONVERT(Varchar(2),ISNULL((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/60)%60)))%60,00)) as totalTask " +
                  "from RTM_Records, RTM_SubTask_List where R_SubTask = STL_ID and R_User_Name='" + user + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_Start_Date_Time))) ='" + indianTime.AddDays(-1).ToShortDateString() + "' and R_Duration != 'HH:MM:SS' and STL_SubTask NOT Like 'Peer Support%'", con);
            da.Fill(ds, "task");
            return ds;
        }

        private DataSet GetActualMeetingDuration(string user)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("meeting"))
            {
                ds.Tables.Remove(ds.Tables["meeting"]);
            }

            da = new SqlDataAdapter("select  CONVERT(Varchar(2),ISNULL(sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/3600,00)) +':'+ CONVERT(Varchar(2),ISNULL((sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/60)%60,00))+':'+CONVERT(Varchar(2),ISNULL((sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/60)%60)))%60,00)) as [Meeting] " +
                            "from RTM_Log_Actions where LA_User_Name='" + user + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_Start_Date_Time)))= '" + indianTime.AddDays(-1).ToShortDateString() + "' and LA_Reason = 'Meeting'", con);
            da.Fill(ds, "meeting");
            return ds;
        }

        private DataSet GetActualCallDuration(string user)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("call"))
            {
                ds.Tables.Remove(ds.Tables["call"]);
            }

            da = new SqlDataAdapter("select  CONVERT(Varchar(2),ISNULL(sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/3600,00)) +':'+ CONVERT(Varchar(2),ISNULL((sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/60)%60,00))+':'+CONVERT(Varchar(2),ISNULL((sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/60)%60)))%60,00)) as [ConferenceCall] " +
                           "from RTM_Log_Actions where LA_User_Name='" + user + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_Start_Date_Time)))= '" + indianTime.AddDays(-1).ToShortDateString() + "' and LA_Reason = 'Conference Call'", con);
            da.Fill(ds, "call");
            return ds;
        }

        private DataSet GetActualPeerSupportDuration(string user)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("peersupport"))
            {
                ds.Tables.Remove(ds.Tables["peersupport"]);
            }

            da = new SqlDataAdapter("select CONVERT(Varchar(2),ISNULL(sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600,00)) +':'+ CONVERT(Varchar(2),ISNULL((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/60)%60,00))+':'+CONVERT(Varchar(2),ISNULL((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/60)%60)))%60,00)) as [PeerSupport] " +
                              "from RTM_Records, RTM_SubTask_List where R_SubTask = STL_ID and R_User_Name='" + user + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_Start_Date_Time))) ='" + indianTime.AddDays(-1).ToShortDateString() + "' and R_Duration != 'HH:MM:SS' and STL_SubTask Like 'Peer Support%'", con);
            da.Fill(ds, "peersupport");
            return ds;
        }

        private void DisplayEstCompData(int TID)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (indianTime.AddDays(-1).DayOfWeek == DayOfWeek.Saturday)
            {
                return;
            }
            if (indianTime.AddDays(-1).DayOfWeek == DayOfWeek.Sunday)
            {
                return;
            }
            DataRow dr;

            BuidTable();

            ds = GetRTMRecords(TID);
            if (ds.Tables["Records"].Rows.Count > 0)
            {
                foreach (DataRow dr1 in ds.Tables["Records"].Rows)
                {
                    dr = dtResult.NewRow();

                    dr["Employee Name"] = dr1["R_User_Name"];
                    dr["Date"] = indianTime.AddDays(-1).ToShortDateString();
                    dr["Client"] = dr1["CL_ClientName"];
                    dr["Task"] = dr1["TL_Task"];
                    dr["Subtask"] = dr1["STL_SubTask"];
                    dr["RTM Time"] = dr1["hour"] + ":" + dr1["minute"] + ":" + dr1["seconds"];
                    dt = GetEstimateRecords(dr1["R_User_Name"].ToString(), Convert.ToInt32(dr1["R_Client"]), Convert.ToInt32(dr1["R_Task"]), Convert.ToInt32(dr1["R_SubTask"]));

                    if (dt.Rows.Count > 0 && dt.Rows[0]["hour"].ToString().Length > 0)
                    {
                        dr["User Estimate"] = dt.Rows[0]["hour"].ToString() + ":" + dt.Rows[0]["minute"].ToString() + ":" + dt.Rows[0]["seconds"].ToString();
                    }

                    dtResult.Rows.Add(dr);
                }
            }

            ds = GetEstNotInRec(TID);

            if (ds.Tables["onlyest"].Rows.Count > 0)
            {
                foreach (DataRow dr2 in ds.Tables["onlyest"].Rows)
                {
                    dr = dtResult.NewRow();

                    dr["Employee Name"] = dr2["EST_UserName"];
                    dr["Date"] = indianTime.AddDays(-1).ToShortDateString();
                    dr["Client"] = dr2["CL_ClientName"];
                    dr["Task"] = dr2["TL_Task"];
                    dr["Subtask"] = dr2["STL_SubTask"];
                    dr["User Estimate"] = dr2["EST_Duration"];
                    dr["RTM Time"] = "";

                    dtResult.Rows.Add(dr);
                }
            }

            dt = GetLogs(TID);
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr2 in dt.Rows)
                {
                    dr = dtResult.NewRow();
                    dr["Employee Name"] = dr2["LA_User_Name"];
                    dr["Date"] = indianTime.AddDays(-1).ToShortDateString();
                    dr["Client"] = "";
                    dr["Task"] = dr2["LA_Reason"];
                    dr["Subtask"] = "";
                    dr["User Estimate"] = "";
                    dr["RTM Time"] = dr2["hour"] + ":" + dr2["minute"] + ":" + dr2["seconds"];

                    dtResult.Rows.Add(dr);
                }
            }

            if (dtResult.Rows.Count > 0)
            {
                string filePath = "";
                if (TID == 10)
                {
                    filePath = "Resource Utilization Estimate Comparison – Onboarding - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv";
                }
                else if (TID == 13)
                {
                    filePath = "Resource Utilization Estimate Comparison – Implementation - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv";
                }
                else if (TID == 31)
                {
                    filePath = "Resource Utilization Estimate Comparison – VTM Testing - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv";
                }
                else if (TID == 9)
                {
                    filePath = "Resource Utilization Estimate Comparison – Quality Check - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv";
                }
                else if (TID == 11)
                {
                    filePath = "Resource Utilization Estimate Comparison – Audit & Optimize - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv";
                }
                else if (TID == 18)
                {
                    filePath = "Resource Utilization Estimate Comparison – Ops Support - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv";
                }
                else if (TID == 22)
                {
                    filePath = "Resource Utilization Estimate Comparison – Client Services 1 - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv";
                }
                else if (TID == 23)
                {
                    filePath = "Resource Utilization Estimate Comparison – Client Services 2 - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv";
                }
                else if (TID == 29)
                {
                    filePath = "Resource Utilization Estimate Comparison – Invoices - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv";
                }

                //string filePath = "" + Directory.GetCurrentDirectory() + "\\Resource Utilization Estimate Comparison – Onboarding – " + DateTime.Now.AddDays(-1).ToShortDateString() + ".csv";
                //CSVUtility.ToCSV(dtResult, filePath);
                MailMessage message1 = new MailMessage();
                DirectCSV csv = new DirectCSV();
                var data = csv.ExportToCSV(dtResult);
                var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(data);
                MemoryStream ms = new MemoryStream(bytes);
                Attachment attachFile = new Attachment(ms, filePath, "application/csv");
                message1.Attachments.Add(attachFile);

                DataRow dr1;
                BuildESTCMPBodyTable();
                TimeSpan total = TimeSpan.Parse("00:00:00");
                TimeSpan lessThan = TimeSpan.Parse("00:00:00");
                int serialNo = 0;

                ds = GetEstimationsForMailBody(TID);

                if (ds.Tables["estimation"].Rows.Count > 0)
                {
                    foreach (DataRow dr3 in ds.Tables["estimation"].Rows)
                    {
                        dr1 = dtResult.NewRow();
                        serialNo = serialNo + 1;
                        dr1["S.No"] = serialNo;
                        dr1["Employee Name"] = dr3["Employee Name"].ToString();
                        dr1["Date"] = dr3["Date"].ToString();
                        dr1["User Estimate"] = dr3["User Estimate"].ToString();

                        ds = GetActualTaskDuration(dr3["Employee Name"].ToString());
                        if (ds.Tables["task"].Rows.Count > 0)
                        {

                            dr1["RTM Task Time"] = ds.Tables["task"].Rows[0]["totalTask"].ToString();
                        }
                        else
                        {
                            dr1["RTM Task Time"] = "00:00:00";
                        }

                        ds = GetActualMeetingDuration(dr3["Employee Name"].ToString());
                        if (ds.Tables["meeting"].Rows.Count > 0)
                        {
                            dr1["Meeting"] = ds.Tables["meeting"].Rows[0]["Meeting"].ToString();
                        }
                        else
                        {
                            dr1["Meeting"] = "00:00:00";
                        }

                        ds = GetActualCallDuration(dr3["Employee Name"].ToString());
                        if (ds.Tables["call"].Rows.Count > 0)
                        {
                            dr1["Conference Call"] = ds.Tables["call"].Rows[0]["ConferenceCall"].ToString();
                        }
                        else
                        {
                            dr1["Conference Call"] = "00:00:00";
                        }

                        ds = GetActualPeerSupportDuration(dr3["Employee Name"].ToString());
                        if (ds.Tables["peersupport"].Rows.Count > 0)
                        {
                            dr1["Peer Support"] = ds.Tables["peersupport"].Rows[0]["PeerSupport"].ToString();
                        }
                        else
                        {
                            dr1["Peer Support"] = "00:00:00";
                        }

                        TimeSpan totalWork = TimeSpan.Parse(dr1["RTM Task Time"].ToString()).Add(TimeSpan.Parse(dr1["Meeting"].ToString())).Add(TimeSpan.Parse(dr1["Conference Call"].ToString())).Add(TimeSpan.Parse(dr1["Peer Support"].ToString()));

                        if (totalWork > TimeSpan.Parse("08:00:00"))
                        {
                            dr1["Less than 8"] = "00:00:00";
                        }
                        else
                        {
                            dr1["Less than 8"] = TimeSpan.Parse("08:00:00").Subtract(totalWork);
                            total = total.Add(TimeSpan.Parse(dr1["Less than 8"].ToString()));
                        }

                        dtResult.Rows.Add(dr1);
                    }

                    dr1 = dtResult.NewRow();
                    dr1["S.No"] = "";
                    dr1["Employee Name"] = "";
                    dr1["Date"] = "";
                    dr1["User Estimate"] = "";
                    dr1["RTM Task Time"] = "";
                    dr1["Meeting"] = "";
                    dr1["Conference Call"] = "";
                    dr1["Peer Support"] = "TOTAL";
                    dr1["Less than 8"] = total;
                    dtResult.Rows.Add(dr1);


                }

                getDelayHTML(dtResult);

                StringBuilder sb = new StringBuilder();

                sb.AppendLine("Hi All,");
                sb.AppendLine("");
                sb.AppendLine("Please find the attached report for today's Resource Utilization Estimate Comparison.");
                sb.AppendLine("");
                sb.AppendLine("");
                sb.AppendLine(myBuilder.ToString());
                sb.AppendLine("");//here I want the data to       display in table format
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                sb.AppendLine("");


                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);

                if (TID == 10)
                {
                    foreach (string item in OnBoardToAddress)
                    {
                        message1.To.Add(new MailAddress(item));
                    }
                    message1.Subject = "Estimate Comparison Report - Onboarding Team Resource Utilization – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";
                    //message1.Body = "Hi All, " + Environment.NewLine + "Please find the attached report for today's Resource Utilization Estimate Comparison." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.";

                }
                else if (TID == 13)
                {
                    message1.To.Add(new MailAddress("Balaji.Nagabhushanam@tangoe.com"));
                    message1.To.Add(new MailAddress("Nagabharana.Sathyanarayana@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Estimate Comparison Report - Implementation Team Resource Utilization – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";
                    // message1.Body = "Hi All, " + Environment.NewLine + "Please find the attached report for today's Resource Utilization Estimate Comparison." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.";

                }
                else if (TID == 31)
                {
                    message1.To.Add(new MailAddress("Sleema.Joseph@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Estimate Comparison Report - VTM Testing Team Resource Utilization – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";
                    // message1.Body = "Hi All, " + Environment.NewLine + "Please find the attached report for today's Resource Utilization Estimate Comparison." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.";

                }
                else if (TID == 9)
                {
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Johwessly.Chennaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Estimate Comparison Report - Quality Check Team Resource Utilization – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";
                    //message1.Body = "Hi All, " + Environment.NewLine + "Please find the attached report for today's Resource Utilization Estimate Comparison." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.";

                }
                else if (TID == 11)
                {
                    message1.To.Add(new MailAddress("Vinith.bekal@tangoe.com"));
                    message1.To.Add(new MailAddress("Balaji.Goalla@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Estimate Comparison Report - Audit & Optimize Team Resource Utilization – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";
                    // message1.Body = "Hi All, " + Environment.NewLine + "Please find the attached report for today's Resource Utilization Estimate Comparison." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.";

                }
                else if (TID == 18)
                {
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Estimate Comparison Report - Ops Support Team Resource Utilization – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";
                }
                else if (TID == 22)
                {
                    message1.To.Add(new MailAddress("Vikas.Vyas@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Estimate Comparison Report - Client Services 1 Team Resource Utilization – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";
                }
                else if (TID == 23)
                {
                    message1.To.Add(new MailAddress("Monnappa.Badumanda@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Estimate Comparison Report - Client Services 2 Team Resource Utilization – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";
                }
                else if (TID == 29)
                {
                    message1.To.Add(new MailAddress("Vikas.Vyas@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Estimate Comparison Report - Invoices Team Resource Utilization – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";
                }

                message1.Body = sb.ToString();
                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment(filePath);
                //message1.Attachments.Add(attachment);
                message1.IsBodyHtml = true;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                //smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);

                //if (File.Exists(filePath))
                //{
                //    File.Delete(filePath);
                //}
            }
            else
            {
                MailMessage message1 = new MailMessage();
                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);
                if (TID == 10)
                {
                    foreach (string item in OnBoardToAddress)
                    {
                        message1.To.Add(new MailAddress(item));
                    }
                    message1.Subject = "Onboarding Team Resource Utilization Estimate Comparison Report – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";
                }
                else if (TID == 13)
                {
                    message1.To.Add(new MailAddress("Balaji.Nagabhushanam@tangoe.com"));
                    message1.To.Add(new MailAddress("Nagabharana.Sathyanarayana@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Estimate Comparison Report - Implementation Team Resource Utilization – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";
                }
                else if (TID == 31)
                {
                    message1.To.Add(new MailAddress("Sleema.Joseph@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Estimate Comparison Report - VTM Testing Team Resource Utilization – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";
                }
                else if (TID == 9)
                {
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Johwessly.Chennaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Estimate Comparison Report - Quality Check Team Resource Utilization – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";

                }
                else if (TID == 11)
                {
                    message1.To.Add(new MailAddress("Vinith.bekal@tangoe.com"));
                    message1.To.Add(new MailAddress("Balaji.Goalla@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Estimate Comparison Report - Audit & Optimize Team Resource Utilization – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";
                }
                else if (TID == 18)
                {
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Estimate Comparison Report - Ops Support Team Resource Utilization – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";
                }
                else if (TID == 22)
                {
                    message1.To.Add(new MailAddress("Vikas.Vyas@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Estimate Comparison Report - Client Services 1 Team Resource Utilization – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";
                }
                else if (TID == 23)
                {
                    message1.To.Add(new MailAddress("Monnappa.Badumanda@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Estimate Comparison Report - Client Services 2 Team Resource Utilization – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";
                }
                else if (TID == 29)
                {
                    message1.To.Add(new MailAddress("Vikas.Vyas@tangoe.com"));
                    message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                    message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                    message1.Subject = "Estimate Comparison Report - Invoices Team Resource Utilization – ( " + indianTime.AddDays(-1).ToShortDateString() + ")";
                }

                //message1.Subject = "Team Resource Utilization Estimate Report – ( " + DateTime.Now.ToShortDateString() + ")";
                message1.Body = "Hi All, " + Environment.NewLine + "No estimate comparison records found for the day." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.";

                message1.IsBodyHtml = false;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                //smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
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

        private DataSet getUserCountFromEstimateTable(int TID)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("count1"))
            {
                ds.Tables.Remove(ds.Tables["count1"]);
            }
            da = new SqlDataAdapter("select DISTINCT EST_UserName from RTM_Estimation where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, EST_Date)))='" + indianTime.ToShortDateString() + "' and EST_TeamId='" + TID + "'", con);
            da.Fill(ds, "count1");
            return ds;
        }

        private DataSet getUserCountFromUserTable(int TID)
        {
            if (ds.Tables.Contains("count2"))
            {
                ds.Tables.Remove(ds.Tables["count2"]);
            }
            da = new SqlDataAdapter("select UL_User_Name from RTM_User_List left join RTM_Access_Level on UL_Employee_Id=AL_EmployeeId where UL_Team_Id='" + TID + "' and AL_AccessLevel=4 and UL_User_Status = 1", con);
            da.Fill(ds, "count2");
            return ds;
        }

        private void CheckEstUsers(int TID)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (indianTime.DayOfWeek == DayOfWeek.Saturday)
            {
                return;
            }
            if (indianTime.DayOfWeek == DayOfWeek.Sunday)
            {
                return;
            }
            ds = getUserCountFromEstimateTable(TID);
            ds = getUserCountFromUserTable(TID);

            if (ds.Tables["count1"].Rows.Count < ds.Tables["count2"].Rows.Count)
            {
                dt = new DataTable();
                da = new SqlDataAdapter("select UL_User_Name as [User Name] from RTM_User_List left join RTM_Access_Level on UL_Employee_Id=AL_EmployeeId where UL_Team_Id='" + TID + "' and AL_AccessLevel=4 and UL_User_Status = 1 and UL_User_Name NOT IN (select DISTINCT EST_UserName from RTM_Estimation where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, EST_Date)))='" + indianTime.ToShortDateString() + "' and EST_TeamId='" + TID + "') ", con);
                da.Fill(dt);

                //string filePath = "" + Directory.GetCurrentDirectory() + "\\Users – " + DateTime.Now.ToString("MM-dd-yyyy") + ".csv";
                //CSVUtility.ToCSV(dt, filePath);

                getDelayHTML(dt);

                StringBuilder sb = new StringBuilder();



                MailMessage message1 = new MailMessage();
                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);

                if (TID == 10)
                {
                    message1.To.Add(new MailAddress("SMS-LSPOnBoarding@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));

                    message1.Subject = "Onboarding - Utilization Estimate Pending";
                    //message1.Body = "Hi Team, " + Environment.NewLine + "Few users have not yet updated their Resource Utilization Estimate in RTM. Please do so before 3.30 PM." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues." + Environment.NewLine + "Thanks, " + Environment.NewLine + "RTM Support.";
                    sb.AppendLine("Hi Team,");
                    sb.AppendLine("");
                    sb.AppendLine("Few users have not yet updated their Resource Utilization Estimate in RTM. Please do so before 3.30 PM.");
                    sb.AppendLine("");
                    sb.AppendLine("");
                    sb.AppendLine(myBuilder.ToString());
                    sb.AppendLine("");//here I want the data to       display in table format
                    sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                    sb.AppendLine("");

                    message1.Body = sb.ToString();
                }
                else if (TID == 13)
                {
                    message1.To.Add(new MailAddress("Bangalore-Implementation@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));

                    message1.Subject = "Implementation - Utilization Estimate Pending";
                    //message1.Body = "Hi Team, " + Environment.NewLine + "Few users have not yet updated their Resource Utilization Estimate in RTM. Please do so before 1.00 PM." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues." + Environment.NewLine + "Thanks, " + Environment.NewLine + "RTM Support.";

                    sb.AppendLine("Hi Team,");
                    sb.AppendLine("");
                    sb.AppendLine("Few users have not yet updated their Resource Utilization Estimate in RTM. Please do so before 1.00 PM.");
                    sb.AppendLine("");
                    sb.AppendLine("");
                    sb.AppendLine(myBuilder.ToString());
                    sb.AppendLine("");//here I want the data to       display in table format
                    sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                    sb.AppendLine("");

                    message1.Body = sb.ToString();
                }
                else if (TID == 31)
                {
                    message1.To.Add(new MailAddress("Abinesh.Natarajan@tangoe.com"));
                    message1.To.Add(new MailAddress("Naveen.Pillanna@tangoe.com"));
                    message1.To.Add(new MailAddress("vandana.singh@tangoe.com"));
                    message1.To.Add(new MailAddress("Sleema.Joseph@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));

                    message1.Subject = "VTM Testing - Utilization Estimate Pending";
                    //message1.Body = "Hi Team, " + Environment.NewLine + "Few users have not yet updated their Resource Utilization Estimate in RTM. Please do so before 1.00 PM." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues." + Environment.NewLine + "Thanks, " + Environment.NewLine + "RTM Support.";

                    sb.AppendLine("Hi Team,");
                    sb.AppendLine("");
                    sb.AppendLine("Few users have not yet updated their Resource Utilization Estimate in RTM. Please do so before 1.00 PM.");
                    sb.AppendLine("");
                    sb.AppendLine("");
                    sb.AppendLine(myBuilder.ToString());
                    sb.AppendLine("");//here I want the data to       display in table format
                    sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                    sb.AppendLine("");

                    message1.Body = sb.ToString();
                }
                else if (TID == 9)
                {
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Johwessly.Chennaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));

                    message1.Subject = "Quality Check - Utilization Estimate Pending";
                    //message1.Body = "Hi, " + Environment.NewLine + "Resource Utilization Estimate for the users listed are not yet updated." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues." + Environment.NewLine + "Thanks, " + Environment.NewLine + "RTM Support.";

                    sb.AppendLine("Hi,");
                    sb.AppendLine("");
                    sb.AppendLine("Resource Utilization Estimate for the users listed are not yet updated.");
                    sb.AppendLine("");
                    sb.AppendLine("");
                    sb.AppendLine(myBuilder.ToString());
                    sb.AppendLine("");//here I want the data to       display in table format
                    sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                    sb.AppendLine("");

                    message1.Body = sb.ToString();
                }
                else if (TID == 11)
                {
                    message1.To.Add(new MailAddress("Vinith.bekal@tangoe.com"));
                    message1.To.Add(new MailAddress("Balaji.Goalla@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));

                    message1.Subject = "Audit & Optimize - Utilization Estimate Pending";
                    //message1.Body = "Hi, " + Environment.NewLine + "Resource Utilization Estimate for the users listed are not yet updated." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues." + Environment.NewLine + "Thanks, " + Environment.NewLine + "RTM Support.";

                    sb.AppendLine("Hi,");
                    sb.AppendLine("");
                    sb.AppendLine("Resource Utilization Estimate for the users listed are not yet updated.");
                    sb.AppendLine("");
                    sb.AppendLine("");
                    sb.AppendLine(myBuilder.ToString());
                    sb.AppendLine("");//here I want the data to       display in table format
                    sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                    sb.AppendLine("");

                    message1.Body = sb.ToString();
                }
                else if (TID == 18)
                {
                    message1.To.Add(new MailAddress("BLR.OperationsSupport@tangoe.com"));
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));

                    message1.Subject = "Ops Support - Utilization Estimate Pending";
                    //message1.Body = "Hi, " + Environment.NewLine + "Resource Utilization Estimate for the users listed are not yet updated." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues." + Environment.NewLine + "Thanks, " + Environment.NewLine + "RTM Support.";

                    sb.AppendLine("Hi,");
                    sb.AppendLine("");
                    sb.AppendLine("Few users have not yet updated their Resource Utilization Estimate in RTM. Please do so before 10 AM.");
                    sb.AppendLine("");
                    sb.AppendLine("");
                    sb.AppendLine(myBuilder.ToString());
                    sb.AppendLine("");//here I want the data to       display in table format
                    sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                    sb.AppendLine("");

                    message1.Body = sb.ToString();
                }
                else if (TID == 22)
                {
                    //message1.To.Add(new MailAddress("Vikas.Vyas@tangoe.com"));


                    foreach (string item in cs1ToAddress)
                    {
                        message1.To.Add(new MailAddress(item));
                    }

                    message1.To.Add(new MailAddress("Shakuntala.Mariyamballi@tangoe.com"));
                    message1.To.Add(new MailAddress("Pulugundla.Karthik@tangoe.com"));

                    message1.Subject = "Client Services 1 - Utilization Estimate Pending";
                    //message1.Body = "Hi Team, " + Environment.NewLine + "Few users have not yet updated their Resource Utilization Estimate in RTM. Please do so before 1.00 PM." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues." + Environment.NewLine + "Thanks, " + Environment.NewLine + "RTM Support.";

                    sb.AppendLine("Hi Team,");
                    sb.AppendLine("");
                    sb.AppendLine("Few users have not yet updated their Resource Utilization Estimate in RTM. Please do so before 11.30 AM.");
                    sb.AppendLine("");
                    sb.AppendLine("");
                    sb.AppendLine(myBuilder.ToString());
                    sb.AppendLine("");//here I want the data to       display in table format
                    sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                    sb.AppendLine("");

                    message1.Body = sb.ToString();
                }
                else if (TID == 23)
                {
                    foreach (string item in cs2ToAddress)
                    {
                        message1.To.Add(new MailAddress(item));
                    }

                    message1.Subject = "Client Services 2 - Utilization Estimate Pending";
                    //message1.Body = "Hi Team, " + Environment.NewLine + "Few users have not yet updated their Resource Utilization Estimate in RTM. Please do so before 1.00 PM." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues." + Environment.NewLine + "Thanks, " + Environment.NewLine + "RTM Support.";

                    sb.AppendLine("Hi Team,");
                    sb.AppendLine("");
                    sb.AppendLine("Few users have not yet updated their Resource Utilization Estimate in RTM. Please do so before 11.30 AM.");
                    sb.AppendLine("");
                    sb.AppendLine("");
                    sb.AppendLine(myBuilder.ToString());
                    sb.AppendLine("");//here I want the data to       display in table format
                    sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                    sb.AppendLine("");

                    message1.Body = sb.ToString();
                }
                else if (TID == 29)
                {
                    message1.To.Add(new MailAddress("Bangalore.Invoices2@tangoe.com"));
                    message1.To.Add(new MailAddress("Vikas.Vyas@tangoe.com"));
                    message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));

                    message1.Subject = "Invoices - Utilization Estimate Pending";

                    sb.AppendLine("Hi Team,");
                    sb.AppendLine("");
                    sb.AppendLine("Few users have not yet updated their Resource Utilization Estimate in RTM. Please do so before 11.30 AM.");
                    sb.AppendLine("");
                    sb.AppendLine("");
                    sb.AppendLine(myBuilder.ToString());
                    sb.AppendLine("");//here I want the data to       display in table format
                    sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                    sb.AppendLine("");

                    message1.Body = sb.ToString();
                }


                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment(filePath);
                //message1.Attachments.Add(attachment);

                message1.IsBodyHtml = true;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                //smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }
        }

        private DataSet GetDayWiseTotal(int teamId, string uname, string day)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("day"))
            {
                ds.Tables.Remove(ds.Tables["day"]);
            }
            da = new SqlDataAdapter("SELECT  " +
                "sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600 as hour, " +
                "(sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/60)%60 as minute," +
                "(sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/60)%60)))%60 as seconds " +
                "from RTM_Records , RTM_SubTask_List where R_SubTask = STL_ID and STL_SubTask Like 'Peer Support%'and R_TeamId='" + teamId + "' and R_User_Name='" + uname + "' and R_Duration != 'HH:MM:SS' and datename(dw,CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_Start_Date_Time)))) = '" + day + "' and R_Start_Date_Time BETWEEN '" + indianTime.AddDays(-(int)DateTime.Now.DayOfWeek - 6).ToShortDateString() + "'  AND '" + indianTime.AddDays(-(int)DateTime.Now.DayOfWeek).ToShortDateString() + "'", con);
            da.Fill(ds, "day");
            return ds;
        }

        private DataSet GetAllUsers()
        {
            if (ds.Tables.Contains("users"))
            {
                ds.Tables.Remove(ds.Tables["users"]);
            }

            da = new SqlDataAdapter("select UL_User_Name, T_TeamName, UL_Team_Id from RTM_User_List, RTM_Team_List, RTM_Access_Level where UL_Team_Id = T_ID and UL_Employee_Id = AL_EmployeeId and UL_User_Status=1 and AL_AccessLevel=4 order by T_TeamName", con);
            da.Fill(ds, "users");
            return ds;
        }

        private void BuildDayWisePeerSupportTable()
        {
            try
            {
                dtResult = new DataTable();
                DataColumn dc;

                dc = new DataColumn("User Name");
                dtResult.Columns.Add(dc);

                dc = new DataColumn("Team");
                dtResult.Columns.Add(dc);

                dc = new DataColumn("Monday");
                dtResult.Columns.Add(dc);

                dc = new DataColumn("Tuesday");
                dtResult.Columns.Add(dc);

                dc = new DataColumn("Wednesday");
                dtResult.Columns.Add(dc);

                dc = new DataColumn("Thursday");
                dtResult.Columns.Add(dc);

                dc = new DataColumn("Friday");
                dtResult.Columns.Add(dc);

                dc = new DataColumn("Saturday");
                dtResult.Columns.Add(dc);

                dc = new DataColumn("Total");
                dtResult.Columns.Add(dc);
            }
            catch (Exception)
            {


            }
        }

        private void WeeklyPeerSupport()
        {
            try
            {
                DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
                DataRow dr;
                string uname;
                string tname;
                int teamid;
                BuildDayWisePeerSupportTable();

                ds = GetAllUsers();

                if (ds.Tables["users"].Rows.Count > 0)
                {
                    foreach (DataRow dr1 in ds.Tables["users"].Rows)
                    {
                        dr = dtResult.NewRow();


                        uname = dr1["UL_User_Name"].ToString();
                        tname = dr1["T_TeamName"].ToString();
                        teamid = Convert.ToInt32(dr1["UL_Team_Id"]);
                        dr["User Name"] = uname;
                        dr["Team"] = tname;
                        ds = GetDayWiseTotal(teamid, uname, "Monday");

                        if (ds.Tables["day"].Rows.Count > 0 && ds.Tables["day"].Rows[0]["hour"].ToString().Length > 0)
                        {
                            dr["Monday"] = ds.Tables["day"].Rows[0]["hour"] + ":" + ds.Tables["day"].Rows[0]["minute"] + ":" + ds.Tables["day"].Rows[0]["seconds"];
                        }
                        else
                        {
                            dr["Monday"] = "00:00:00";
                        }

                        ds = GetDayWiseTotal(teamid, uname, "Tuesday");

                        if (ds.Tables["day"].Rows.Count > 0 && ds.Tables["day"].Rows[0]["hour"].ToString().Length > 0)
                        {
                            dr["Tuesday"] = ds.Tables["day"].Rows[0]["hour"] + ":" + ds.Tables["day"].Rows[0]["minute"] + ":" + ds.Tables["day"].Rows[0]["seconds"];
                        }
                        else
                        {
                            dr["Tuesday"] = "00:00:00";
                        }

                        ds = GetDayWiseTotal(teamid, uname, "Wednesday");

                        if (ds.Tables["day"].Rows.Count > 0 && ds.Tables["day"].Rows[0]["hour"].ToString().Length > 0)
                        {
                            dr["Wednesday"] = ds.Tables["day"].Rows[0]["hour"] + ":" + ds.Tables["day"].Rows[0]["minute"] + ":" + ds.Tables["day"].Rows[0]["seconds"];
                        }
                        else
                        {
                            dr["Wednesday"] = "00:00:00";
                        }

                        ds = GetDayWiseTotal(teamid, uname, "Thursday");

                        if (ds.Tables["day"].Rows.Count > 0 && ds.Tables["day"].Rows[0]["hour"].ToString().Length > 0)
                        {
                            dr["Thursday"] = ds.Tables["day"].Rows[0]["hour"] + ":" + ds.Tables["day"].Rows[0]["minute"] + ":" + ds.Tables["day"].Rows[0]["seconds"];
                        }
                        else
                        {
                            dr["Thursday"] = "00:00:00";
                        }

                        ds = GetDayWiseTotal(teamid, uname, "Friday");

                        if (ds.Tables["day"].Rows.Count > 0 && ds.Tables["day"].Rows[0]["hour"].ToString().Length > 0)
                        {
                            dr["Friday"] = ds.Tables["day"].Rows[0]["hour"] + ":" + ds.Tables["day"].Rows[0]["minute"] + ":" + ds.Tables["day"].Rows[0]["seconds"];
                        }
                        else
                        {
                            dr["Friday"] = "00:00:00";
                        }

                        ds = GetDayWiseTotal(teamid, uname, "Saturday");

                        if (ds.Tables["day"].Rows.Count > 0 && ds.Tables["day"].Rows[0]["hour"].ToString().Length > 0)
                        {
                            dr["Saturday"] = ds.Tables["day"].Rows[0]["hour"] + ":" + ds.Tables["day"].Rows[0]["minute"] + ":" + ds.Tables["day"].Rows[0]["seconds"];
                        }
                        else
                        {
                            dr["Saturday"] = "00:00:00";
                        }

                        TimeSpan total = TimeSpan.Parse(dr["Monday"].ToString()) + TimeSpan.Parse(dr["Tuesday"].ToString()) + TimeSpan.Parse(dr["Wednesday"].ToString()) + TimeSpan.Parse(dr["Thursday"].ToString()) + TimeSpan.Parse(dr["Friday"].ToString()) + TimeSpan.Parse(dr["Saturday"].ToString());

                        dr["Total"] = total;

                        dtResult.Rows.Add(dr);
                    }

                    if (dtResult.Rows.Count > 0)
                    {
                        MailMessage message1 = new MailMessage();
                        string filePath = "Time Spent on Peer Support for Week – " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToString("MM-dd-yyyy") + "-" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToString("MM-dd-yyyy") + ".csv";
                        //CSVUtility.ToCSV(dtResult, filePath);

                        DirectCSV csv = new DirectCSV();
                        var data = csv.ExportToCSV(dtResult);
                        var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(data);
                        MemoryStream ms = new MemoryStream(bytes);

                        Attachment attachFile = new Attachment(ms, filePath, "application/csv");
                        message1.Attachments.Add(attachFile);

                        StringBuilder sb = new StringBuilder();
                        try
                        {
                            BuildPeerSupportBody();

                            ds = GetTeams();
                            if (ds.Tables["teams"].Rows.Count > 0)
                            {
                                DataRow dr1;

                                foreach (DataRow dr2 in ds.Tables["teams"].Rows)
                                {
                                    dr1 = dtResult.NewRow();
                                    int teamId1 = Convert.ToInt32(dr2["T_ID"]);
                                    int totalHours = -0;
                                    int totalMeetings = 0;
                                    int totalPeerSupport = 0;

                                    dr1["Team"] = dr2["T_TeamName"].ToString();

                                    ds = GetTotalWorkHours(teamId1);
                                    if (ds.Tables["TotalWork"].Rows.Count > 0)
                                    {
                                        totalHours = Convert.ToInt32(ds.Tables["TotalWork"].Rows[0]["hour"]);
                                    }

                                    ds = GetMeetingHours(teamId1);
                                    if (ds.Tables["TotalMeeting"].Rows.Count > 0)
                                    {
                                        totalMeetings = Convert.ToInt32(ds.Tables["TotalMeeting"].Rows[0]["hour"]);
                                    }

                                    dr1["Total Hours Worked"] = totalHours + totalMeetings;

                                    ds = GetTotalPeerSupport(teamId1);

                                    if (ds.Tables["TotalPeer"].Rows.Count > 0)
                                    {
                                        totalPeerSupport = Convert.ToInt32(ds.Tables["TotalPeer"].Rows[0]["hour"]);
                                    }

                                    dr1["Peer Support"] = totalPeerSupport;

                                    int totalWorkHours = totalHours + totalMeetings;
                                    if (totalWorkHours != 0)
                                    {
                                        decimal percent = (Convert.ToDecimal(totalPeerSupport) / Convert.ToDecimal(totalWorkHours)) * 100;
                                        dr1["% of Peer Support"] = decimal.Round(percent, 2, MidpointRounding.AwayFromZero);
                                    }
                                    else
                                    {
                                        dr1["% of Peer Support"] = "0";
                                    }

                                    dtResult.Rows.Add(dr1);
                                }

                            }

                            getDelayHTML(dtResult);



                            sb.AppendLine("Hi All,");
                            sb.AppendLine("");
                            sb.AppendLine("Please find attached the time spent by users on Peer Support last week.");
                            sb.AppendLine("");
                            sb.AppendLine("");
                            sb.AppendLine(myBuilder.ToString());
                            sb.AppendLine("");//here I want the data to       display in table format
                            sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                            sb.AppendLine("");
                        }
                        catch (Exception)
                        {

                        }


                        SmtpClient smtp = new SmtpClient();

                        message1.From = new MailAddress(FromAddress);

                        message1.To.Add(new MailAddress("Bangalore.OpsManagers@tangoe.com"));
                        message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                        message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                        message1.Subject = "Time Spent on Peer Support for Week – ( " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToShortDateString() + "-" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToShortDateString() + ")";
                        //message1.Body = "Hi All, " + Environment.NewLine + "Please find attached the time spent by users on Peer Support last week." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.";
                        message1.Body = sb.ToString();


                        //System.Net.Mail.Attachment attachment;
                        //attachment = new System.Net.Mail.Attachment(filePath);
                        //message1.Attachments.Add(attachment);
                        message1.IsBodyHtml = true;

                        smtp.Port = 25;
                        smtp.Host = "10.0.5.104";
                        //smtp.Host = "outlook-south.tangoe.com";
                        smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                        smtp.EnableSsl = false;

                        smtp.Send(message1);

                        //if (File.Exists(filePath))
                        //{
                        //    File.Delete(filePath);
                        //}
                    }
                }
            }
            catch (Exception)
            {

            }
        }

        private void BuildPeerSupportBody()
        {
            dtResult = new DataTable();

            DataColumn dc;

            dc = new DataColumn("Team");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Total Hours Worked");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("Peer Support");
            dtResult.Columns.Add(dc);

            dc = new DataColumn("% of Peer Support");
            dtResult.Columns.Add(dc);
        }

        private DataSet GetTeams()
        {
            if (ds.Tables.Contains("teams"))
            {
                ds.Tables.Remove(ds.Tables["teams"]);
            }
            da = new SqlDataAdapter("select T_ID, T_TeamName from dbo.RTM_Team_List where T_Active =1 order by T_TeamName", con);
            da.Fill(ds, "teams");
            return ds;
        }

        private DataSet GetTotalWorkHours(int teamId)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("TotalWork"))
            {
                ds.Tables.Remove(ds.Tables["TotalWork"]);
            }
            da = new SqlDataAdapter("select ISNULL(sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600,00) as hour from RTM_Records where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_Start_Date_Time))) Between '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToShortDateString() + "' and '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToShortDateString() + "' and R_TeamId=" + teamId + " and R_Duration != 'HH:MM:SS'", con);
            da.Fill(ds, "TotalWork");
            return ds;
        }

        private DataSet GetMeetingHours(int teamId)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("TotalMeeting"))
            {
                ds.Tables.Remove(ds.Tables["TotalMeeting"]);
            }
            da = new SqlDataAdapter("select ISNULL(sum(datediff(second,'00:00:00',REPLACE(LA_Duration,'-', '')))/3600,00) as hour from RTM_Log_Actions where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, LA_Start_Date_Time))) Between '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToShortDateString() + "' and '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToShortDateString() + "' and LA_TeamId=" + teamId + " and LA_Reason='Meeting' and LA_Duration != 'HH:MM:SS'", con);
            da.Fill(ds, "TotalMeeting");
            return ds;
        }

        private DataSet GetTotalPeerSupport(int teamId)
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            if (ds.Tables.Contains("TotalPeer"))
            {
                ds.Tables.Remove(ds.Tables["TotalPeer"]);
            }
            da = new SqlDataAdapter("select ISNULL(sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600,00) as hour from RTM_Records, RTM_SubTask_List where R_SubTask = STL_ID and STL_SubTask Like 'Peer Support%' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_Start_Date_Time))) Between '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToShortDateString() + "' and '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToShortDateString() + "' and R_TeamId=" + teamId + " and R_Duration != 'HH:MM:SS'", con);
            da.Fill(ds, "TotalPeer");
            return ds;
        }

        private DataTable GetEarlyLogoffDetails()
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            dt = new DataTable();

            da = new SqlDataAdapter("select ROW_NUMBER() over (order by EL_User_Name) as [Sl.No.], EL_User_Name as [User Name], CONVERT(VARCHAR(10), EL_Date, 111) as Date, LTRIM(RIGHT(CONVERT(VARCHAR(20), EL_Scheduled, 100), 7)) as [Scheduled Logoff], EL_Actual as [Actual Logoff], EL_Total_Office_Hours as [Total Office Hours], EL_Reason as Reason, EL_Comments as Comments from RTM_EarlyLogOffDetails where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, EL_Actual))) BETWEEN '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToShortDateString() + "'  AND '" + indianTime.AddDays(-(int)indianTime.DayOfWeek).ToShortDateString() + "' group BY EL_User_Name,EL_Actual, EL_Scheduled, EL_Reason, EL_Comments, EL_Date,EL_Total_Office_Hours  Having ISNULL(sum(datediff(second,'00:00:00',EL_Total_Office_Hours))/3600,00) < 9 order by EL_User_Name", con);
            da.Fill(dt);
            return dt;
        }

        private void SendEarlyLogOffDetails()
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            dt = GetEarlyLogoffDetails();

            if (dt.Rows.Count > 0)
            {
                //string filePath = "" + Directory.GetCurrentDirectory() + "\\Early Logoff Detected for last week – " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToString("MM-dd-yyyy") + "-" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToString("MM-dd-yyyy") + ".csv";
                //CSVUtility.ToCSV(dt, filePath);

                MailMessage message1 = new MailMessage();

                DirectCSV csv = new DirectCSV();
                var data = csv.ExportToCSV(dt);
                var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(data);
                MemoryStream ms = new MemoryStream(bytes);

                Attachment attachFile = new Attachment(ms, "Early Logoff Detected for last week – " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToString("MM-dd-yyyy") + "-" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToString("MM-dd-yyyy") + ".csv", "application/csv");
                message1.Attachments.Add(attachFile);

                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);

                message1.To.Add(new MailAddress("Bangalore.OpsManagers@tangoe.com"));
                message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));
                message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));

                message1.Subject = "Early Logoff Detected for last week – ( " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToShortDateString() + "-" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToShortDateString() + ")";
                message1.Body = "Hi All, " + Environment.NewLine + "Please find attached the early logoff report for last week. Please update your comments in the RTM Reports under Early Logoff report page. You can find the page under Reports => Time Card => Early Logoff" + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.";



                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment(filePath);
                //message1.Attachments.Add(attachment);
                message1.IsBodyHtml = false;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                //smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);

                //if (File.Exists(filePath))
                //{
                //    File.Delete(filePath);
                //}
            }
        }

        //HRIS DATA UPLOAD

        private void UploadHRISData()
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            var sourceLoc = @"D:\Development\TRQ-2016-11-30";
            //var sourceLoc = @"\\files\Shares\HRIS Data\Production";
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
                                string sQuery = "insert into RTM_Master_UserList (MUL_EmployeeId,MUL_FirstName,MUL_LastName,MUL_EmailId,MUL_ManagerID,MUL_CreatedOn,MUL_ActiveStatus) " +
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


        /// <summary>
        /// Uploads data from HRIS to RTM Master user list table 07/11/2018
        /// </summary>
        private void UploadHRISDataNew()
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            var sourceLoc = @"D:\Development\TRQ-2016-11-30";
            //var sourceLoc = @"\\files\Shares\HRIS Data\Production";
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
                        ClearMasterList();
                        //DeactivateAllEmployees();
                        // dtEmp = CheckExistingEmployee();
                        foreach (DataRow drRow in data.Rows)
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
                                    new SqlParameter("@status", 1),
                                    new SqlParameter("@title", drRow["title"].ToString()),
                                    new SqlParameter("@empLocation", drRow["custom_attribute_5"].ToString())
                                };
                            string sQuery = "insert into RTM_Master_UserList (MUL_EmployeeId,MUL_FirstName,MUL_LastName,MUL_EmailId,MUL_ManagerID,MUL_CreatedOn,MUL_ActiveStatus, MUL_Job_Title, MUL_Emp_Location) " +
                                             "values(@empId, @first,@last,@emailId,@managerId,@createdOn,@status,@title, @empLocation)";
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

        private void ClearMasterList()
        {
            using (cmd = new SqlCommand("Truncate table RTM_Master_UserList", globalCon))
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

        private void UpdateUserListTable()
        {
            dt = new DataTable();
            using (SqlDataAdapter da = new SqlDataAdapter("select UL_Employee_Id as [Emp Id], UL_User_Name as [User Name], T_TeamName as [Team Name], UL_RepMgrId as [Old Reporting Manager], MUL_ManagerID as [New Reporting Manager]  from RTM_User_List, RTM_Team_List, RTM_Master_UserList where UL_Team_Id= T_ID and UL_Employee_Id = MUL_EmployeeId and UL_User_Status =1 and UL_RepMgrId != MUL_ManagerID", globalCon))
            {
                da.Fill(dt);
            }

            if (dt.Rows.Count > 0)
            {
                getDelayHTML(dt);

                string sQuery = "Update ul Set ul.UL_RepMgrId = mu.MUL_ManagerID , ul.UL_RepMgrEmail = mu.MUL_ManagerEmail_id From RTM_User_List ul join RTM_Master_UserList mu on UL_Employee_Id = MUL_EmployeeId Where UL_User_Status =1 and UL_RepMgrId != MUL_ManagerID ";

                using (cmd = new SqlCommand(sQuery, globalCon))
                {
                    globalCon.Open();
                    cmd.ExecuteNonQuery();
                    globalCon.Close();
                }
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Hi Team,");

                sb.AppendLine("");
                sb.AppendLine("");
                sb.AppendLine(myBuilder.ToString());
                sb.AppendLine("");
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                sb.AppendLine("");

                MailMessage message1 = new MailMessage();
                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);

                message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));

                message1.Subject = "Users updated with their new reporting manager id";
                //message1.Body = "Hi All, " + Environment.NewLine + "Please find attached the time spent by users on Peer Support last week." + Environment.NewLine + "" + Environment.NewLine + "This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.";
                message1.Body = sb.ToString();


                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment(filePath);
                //message1.Attachments.Add(attachment);
                message1.IsBodyHtml = true;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                //smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }

        }

        private void OffshoreTasks()
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            dt = new DataTable();
            da = new SqlDataAdapter("select ROW_NUMBER() over (order by T_TeamName, STL_Process_Code) as [Sl.No.], T_TeamName as [Team], STL_Process_Code as [Process Code], sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600 as Hours, " +
                            "(sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/60)%60 as Minutes, " +
                            "(sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))-(((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/3600)*3600)-60*((sum(datediff(second,'00:00:00',REPLACE(R_Duration,'-', '')))/60)%60)))%60 as Seconds " +
                            "from RTM_Records, RTM_SubTask_List, RTM_Team_List " +
                            "where R_SubTask = STL_ID and R_TeamId = T_ID and STL_Process_Code is not null " +
                            "and R_Duration != 'HH:MM:SS' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_Start_Date_Time))) " +
                            "BETWEEN '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToShortDateString() + "' and '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToShortDateString() + "' " +
                            "GROUP BY T_TeamName, STL_Process_Code order By T_TeamName, STL_Process_Code", con);
            da.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                getDelayHTML(dt);

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Hi All,");

                sb.AppendLine("Please find below time spent on Offshored tasks for last week.");
                sb.AppendLine("");
                sb.AppendLine(myBuilder.ToString());
                sb.AppendLine("");
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                sb.AppendLine("");

                MailMessage message1 = new MailMessage();
                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);

                //message1.To.Add(new MailAddress("Vinith.Bekal@tangoe.com"));
                message1.To.Add(new MailAddress("Vikas.Vyas@tangoe.com"));
                message1.To.Add(new MailAddress("Sree.Sandirasegarane@tangoe.com"));
                //message1.To.Add(new MailAddress("Rohini.Kulkarni@tangoe.com"));
                //message1.To.Add(new MailAddress("Nandini.Rajagopal@tangoe.com"));
                //message1.To.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                message1.To.Add(new MailAddress("RTM-Support@tangoe.com"));

                message1.Subject = "Time Spent on Offshored tasks(" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToString("MM-dd-yyyy") + "-" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToString("MM-dd-yyyy") + ")";

                message1.Body = sb.ToString();

                message1.IsBodyHtml = true;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                // smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }
        }


        private void incompleteData()
        {
            DataTable dtUsers = new System.Data.DataTable();
            DataTable dtRecords = new DataTable();
            DataTable dtLogActions = new DataTable();
            DataTable dtLeave = new DataTable();
            DataTable dtFilled = new DataTable();
            dtFilled.Clear();
            dtFilled.Columns.Add("Date", typeof(DateTime));
            dtFilled.Columns.Add("Duration", typeof(decimal));


            da = new SqlDataAdapter("SELECT * FROM dbo.RTM_User_List INNER JOIN dbo.RTM_Team_List ON dbo.RTM_User_List.UL_Team_Id = dbo.RTM_Team_List.T_ID"
            + " WHERE (dbo.RTM_Team_List.T_Location = 'IND' or dbo.RTM_Team_List.T_Location = 'CHN' or dbo.RTM_Team_List.T_Location = 'OTH') AND (dbo.RTM_User_List.UL_User_Status = 1) "
            + "and (dbo.RTM_User_List.UL_User_Name not in('cme.dst','cme.ip','Pascale Sadaka','Marcelino Ayoub','Canita Dahdah','Caragh Gravitt')) and (dbo.RTM_User_List.UL_Hourly <> 1 or dbo.RTM_User_List.UL_Hourly is Null) "
            + " and (UL_Exclude <> 1 or UL_Exclude is null) order by dbo.RTM_Team_List.T_Location", con);
            da.Fill(dtUsers);
            //            da = new SqlDataAdapter("SELECT        dbo.RTM_User_List.* "
            //+ "FROM            dbo.RTM_User_List INNER JOIN "
            //                       + "  dbo.RTM_Team_List ON dbo.RTM_User_List.UL_Team_Id = dbo.RTM_Team_List.T_ID "
            //+ " WHERE        (dbo.RTM_Team_List.T_ID = 4)", con);
            //            da.Fill(dtUsers);

            //da = new SqlDataAdapter("SELECT        dbo.RTM_User_List.*FROM            dbo.RTM_User_List WHERE        (UL_User_Name = 'Sleema Joseph')", con);
            //da.Fill(dtUsers);


            foreach (DataRow dr in dtUsers.Rows)
            {
                try
                {
                    dtFilled.Rows.Clear();
                    dtRecords.Clear();
                    dtLogActions.Clear();
                    dtLeave.Clear();
                    da = new SqlDataAdapter("Select  distinct CONVERT(date,R_Start_Date_Time,101) as R_Start_Date_Time ,sum(CAST((datediff(second,0,R_Duration) / (60.0 * 60.0)) AS DECIMAL(10, 4))) as R_Duration from RTM_Records WITH (NOLOCK) "
                   + " where R_User_Name = @username and R_Employee_Id = @empId and R_Status = 'Completed' and R_Submit = '1' and (R_Start_Date_Time >= @startTime and "
                    + "R_Start_Date_Time <= @endTime) group by R_Start_Date_Time,R_Duration order by R_Start_Date_Time", con);

                    da.SelectCommand.Parameters.Clear();
                    da.SelectCommand.Parameters.Add(new SqlParameter
                {
                    ParameterName = "@username",
                    Value = dr["UL_User_Name"]
                });

                    da.SelectCommand.Parameters.Add(new SqlParameter
                    {
                        ParameterName = "@empId",
                        Value = dr["UL_Employee_Id"]
                    });

                    da.SelectCommand.Parameters.Add(new SqlParameter
                    {
                        ParameterName = "@startTime",
                        Value = DateTime.Now.AddDays(-14)
                    });

                    da.SelectCommand.Parameters.Add(new SqlParameter
                    {
                        ParameterName = "@endTime",
                        Value = DateTime.Now
                    });

                    da.Fill(dtRecords);

                    da = new SqlDataAdapter("Select  distinct CONVERT(date,LA_Start_Date_Time,101) as LA_Start_Date_Time ,sum(CAST((datediff(second,0,LA_Duration) / (60.0 * 60.0)) "
                    + "AS DECIMAL(10, 4))) as LA_Duration from RTM_Log_Actions WITH (NOLOCK) where LA_User_Name = @username and  LA_Duration <> '' and LA_Submit = '1'  and "
                    + "(LA_Start_Date_Time >= @startTime and LA_Start_Date_Time <= @endTime) group by LA_Start_Date_Time,LA_Duration order by LA_Start_Date_Time", con);

                    da.SelectCommand.Parameters.Clear();
                    da.SelectCommand.Parameters.Add(new SqlParameter
                    {
                        ParameterName = "@username",
                        Value = dr["UL_User_Name"]
                    });

                    da.SelectCommand.Parameters.Add(new SqlParameter
                    {
                        ParameterName = "@startTime",
                        Value = DateTime.Now.AddDays(-14)
                    });

                    da.SelectCommand.Parameters.Add(new SqlParameter
                    {
                        ParameterName = "@endTime",
                        Value = DateTime.Now
                    });


                    da.Fill(dtLogActions);


                    //da = new SqlDataAdapter("Select  distinct CONVERT(date,LD_Date,101) as LD_Date ,LD_Duration from RTM_LeaveDetails WITH (NOLOCK) where LD_UserName = '" + dr["UL_User_Name"] + "' and LD_Submit = '1' and (LD_Date >= '" + DateTime.Now.AddDays(-14) + "' and LD_Date <= '" + DateTime.Now + "')", con);
                    //da.Fill(dtLeave);

                    foreach (DataRow dr1 in dtRecords.Rows)
                    {
                        dtFilled.Rows.Add(dr1["R_Start_Date_Time"], dr1["R_Duration"]);
                    }

                    foreach (DataRow dr2 in dtLogActions.Rows)
                    {
                        dtFilled.Rows.Add(dr2["LA_Start_Date_Time"], dr2["LA_Duration"]);
                    }

                    //foreach (DataRow dr3 in dtLeave.Rows)
                    //{
                    //    if (Convert.IsDBNull(dr3["LD_Duration"]))
                    //    {
                    //        dtFilled.Rows.Add(dr3["LD_Date"], 8);
                    //    }
                    //    else
                    //    {
                    //        dtFilled.Rows.Add(dr3["LD_Date"], dr3["LD_Duration"]);
                    //    }

                    //}


                    int Count = dtFilled
                                .AsEnumerable()
                                .Select(r => r.Field<DateTime>("Date"))
                                .Distinct()
                                .Count();

                    decimal sum = dtFilled.AsEnumerable().Sum(x => x.Field<decimal>("Duration"));

                    if (Convert.ToInt16(Count) == 0)
                    {
                        if (dr["UL_RepMgrEmail"].ToString() != "")
                        {
                            if (dr["UL_EmailId"].ToString().Trim() != "")
                            {
                                sendmailToUserandManager(dr["UL_User_Name"].ToString(), dr["UL_RepMgrEmail"].ToString().Trim(), dr["UL_EmailId"].ToString().Trim(), Count, sum);
                            }

                        }
                        else
                        {
                            if (dr["UL_EmailId"].ToString().Trim() != "")
                            {
                                sendmailToUserandManager(dr["UL_User_Name"].ToString(), dr["UL_EmailId"].ToString().Trim(), dr["UL_EmailId"].ToString().Trim(), Count, sum);
                            }
                        }

                        continue;
                    }

                    else
                    {

                        dtFilled.Rows.Clear();
                        dtRecords.Clear();
                        dtLogActions.Clear();
                        dtLeave.Clear();

                        da = new SqlDataAdapter("Select  distinct CONVERT(date,R_Start_Date_Time,101) as R_Start_Date_Time ,sum(CAST((datediff(second,0,R_Duration) / (60.0 * 60.0)) AS DECIMAL(10, 4))) as R_Duration from RTM_Records WITH (NOLOCK) "
                       + " where R_User_Name = @username and R_Employee_Id = @empId and R_Status = 'Completed' and R_Submit = '1' and (R_Start_Date_Time >= @startTime and "
                        + "R_Start_Date_Time <= @endTime) group by R_Start_Date_Time,R_Duration order by R_Start_Date_Time", con);


                        da.SelectCommand.Parameters.Clear();
                        da.SelectCommand.Parameters.Add(new SqlParameter
                        {
                            ParameterName = "@username",
                            Value = dr["UL_User_Name"]
                        });

                        da.SelectCommand.Parameters.Add(new SqlParameter
                        {
                            ParameterName = "@empId",
                            Value = dr["UL_Employee_Id"]
                        });

                        da.SelectCommand.Parameters.Add(new SqlParameter
                        {
                            ParameterName = "@startTime",
                            Value = DateTime.Now.AddDays(-7)
                        });

                        da.SelectCommand.Parameters.Add(new SqlParameter
                        {
                            ParameterName = "@endTime",
                            Value = DateTime.Now
                        });

                        da.Fill(dtRecords);


                        da = new SqlDataAdapter("Select  distinct CONVERT(date,LA_Start_Date_Time,101) as LA_Start_Date_Time ,sum(CAST((datediff(second,0,LA_Duration) / (60.0 * 60.0)) "
                        + "AS DECIMAL(10, 4))) as LA_Duration from RTM_Log_Actions WITH (NOLOCK) where LA_User_Name = @username and LA_Duration <> ''  and LA_Submit = '1'  and "
                        + "(LA_Start_Date_Time >= @startTime and LA_Start_Date_Time <= @endTime) group by LA_Start_Date_Time,LA_Duration order by LA_Start_Date_Time", con);

                        da.SelectCommand.Parameters.Clear();
                        da.SelectCommand.Parameters.Add(new SqlParameter
                        {
                            ParameterName = "@username",
                            Value = dr["UL_User_Name"]
                        });

                        da.SelectCommand.Parameters.Add(new SqlParameter
                        {
                            ParameterName = "@startTime",
                            Value = DateTime.Now.AddDays(-7)
                        });

                        da.SelectCommand.Parameters.Add(new SqlParameter
                        {
                            ParameterName = "@endTime",
                            Value = DateTime.Now
                        });


                        da.Fill(dtLogActions);

                        //da = new SqlDataAdapter("Select  distinct CONVERT(date,LD_Date,101) as LD_Date ,LD_Duration from RTM_LeaveDetails WITH (NOLOCK) where LD_UserName = '" + dr["UL_User_Name"] + "' and  LD_Submit = '1' and (LD_Date >= '" + DateTime.Now.AddDays(-7) + "' and LD_Date <= '" + DateTime.Now + "')", con);
                        //da.Fill(dtLeave);

                        foreach (DataRow dr1 in dtRecords.Rows)
                        {
                            dtFilled.Rows.Add(dr1["R_Start_Date_Time"], dr1["R_Duration"]);
                        }

                        foreach (DataRow dr2 in dtLogActions.Rows)
                        {
                            dtFilled.Rows.Add(dr2["LA_Start_Date_Time"], dr2["LA_Duration"]);
                        }

                        //foreach (DataRow dr3 in dtLeave.Rows)
                        //{
                        //    if (Convert.IsDBNull(dr3["LD_Duration"]))
                        //    {
                        //        dtFilled.Rows.Add(dr3["LD_Date"], 8);
                        //    }
                        //        else
                        //    {
                        //        dtFilled.Rows.Add(dr3["LD_Date"], dr3["LD_Duration"]);
                        //    }
                        //}

                        int Count1;
                        Count1 = dtFilled
                                   .AsEnumerable()
                                   .Select(r => r.Field<DateTime>("Date"))
                                   .Distinct()
                                   .Count();

                        decimal sum1 = dtFilled.AsEnumerable().Sum(x => x.Field<decimal>("Duration"));

                        if (Convert.ToInt16(Count1) == 0)
                        {
                            if (dr["UL_EmailId"].ToString().Trim() != "")
                            {
                                sendmailToUser(dr["UL_User_Name"].ToString(), dr["UL_EmailId"].ToString().Trim(), Count1, sum1);
                            }
                        }
                    }
                }

                catch (Exception ex)
                {
                    sendmailToRTMTeam("Mohammed.Shiddique@Tangoe.com", ex.ToString() + dr["UL_User_Name"].ToString());
                    //MessageBox.Show(ex.ToString());
                }
            }
            sendmailToRTMTeam("Mohammed.Shiddique@Tangoe.com", "Reminders completed successfully");
        }

        private void sendmailToUser(string userName, string emailId, int dayCount, decimal hours)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<p>Hi " + userName + ",</p>");
            sb.AppendLine("");
            sb.AppendLine("<p>It looks like you may have missed to submit your time sheet for last week. Please review and submit. Please ignore if already verified. To submit your time sheet now please visit : http://10.55.6.155/RTMGlobal/ </p>");
            sb.AppendLine("");
            //sb.AppendLine("<p>Day count : " + dayCount + ". Hours recorded : " + hours + "</p>");
            //sb.AppendLine(myBuilder.ToString());
            // sb.AppendLine("");
            sb.AppendLine("");
            sb.AppendLine("<p>Regards,</p>");
            //sb.AppendLine("");
            sb.AppendLine("<p><b>RTM Support</b></p>");
            sb.AppendLine("");
            sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
            sb.AppendLine("");
            using (MailMessage message1 = new MailMessage())
            {
                using (SmtpClient smtp = new SmtpClient())
                {
                    message1.From = new MailAddress(FromAddress);
                    //message1.From = new MailAddress("Mohammed.Shiddique@Tangoe.com");

                    message1.To.Add(new MailAddress(emailId));

                    message1.Subject = "RTM Reminder - Time sheet not submitted!";

                    message1.Body = sb.ToString();

                    message1.IsBodyHtml = true;

                    //SmtpClient smtpClient = new SmtpClient("mail.north.tangoe.com");
                    //smtpClient.UseDefaultCredentials = false;
                    //NetworkCredential credentials = new NetworkCredential("Mohammed.Shiddique", "Yunoosjan123!");
                    //smtpClient.Credentials = credentials;

                    //smtpClient.Send(message1);

                    smtp.Port = 25;
                    smtp.Host = "10.0.5.104";
                    //smtp.Host = "outlook-south.tangoe.com";
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    smtp.EnableSsl = false;

                    smtp.Send(message1);
                }
            }
            //MailMessage message1 = new MailMessage();
            //SmtpClient smtp = new SmtpClient();


        }

        private void sendmailToUserandManager(string userName, string managerEmailId, string emailId, int dayCount, decimal hours)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<p>Hi,</p>");
            sb.AppendLine("");
            sb.AppendLine("<p>You may have missed to submit your time sheet for last 2 weeks or more in RTM. To submit your time sheet now please visit : http://10.55.6.155/RTMGlobal/ </p>");
            sb.AppendLine("");
            //sb.AppendLine("<p>Day count : " + dayCount + ". Hours recorded : " + hours + "</p>");
            //sb.AppendLine(myBuilder.ToString());
            sb.AppendLine("");
            sb.AppendLine("<p>Regards,</p>");
            //sb.AppendLine("");
            sb.AppendLine("<p><b>RTM Support</b></p>");
            sb.AppendLine("");
            sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
            sb.AppendLine("");

            using (MailMessage message1 = new MailMessage())
            {
                using (SmtpClient smtp = new SmtpClient())
                {
                    message1.From = new MailAddress(FromAddress);
                    //message1.From = new MailAddress("Mohammed.Shiddique@Tangoe.com");

                    message1.To.Add(new MailAddress(emailId));
                    //message1.CC.Add(new MailAddress(managerEmailId));

                    message1.Subject = "RTM Reminder - Time sheet not submitted!";

                    message1.Body = sb.ToString();

                    message1.IsBodyHtml = true;

                    //SmtpClient smtpClient = new SmtpClient("mail.north.tangoe.com");
                    //smtpClient.UseDefaultCredentials = false;
                    //NetworkCredential credentials = new NetworkCredential("Mohammed.Shiddique", "Yunoosjan123!");
                    //smtpClient.Credentials = credentials;
                    //smtpClient.Send(message1);

                    smtp.Port = 25;
                    smtp.Host = "10.0.5.104";
                    // smtp.Host = "outlook-south.tangoe.com";
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    smtp.EnableSsl = false;
                    smtp.Send(message1);
                }

            }
            //MailMessage message1 = new MailMessage();
            //SmtpClient smtp = new SmtpClient();           
        }

        private void sendmailToRTMTeam(string emailId, string status)
        {
            StringBuilder sb = new StringBuilder();

            MailMessage message1 = new MailMessage();
            SmtpClient smtp = new SmtpClient();

            message1.From = new MailAddress(FromAddress);

            message1.To.Add(new MailAddress(emailId));

            message1.Subject = "RTM Reminder Status";

            message1.Body = status;

            message1.IsBodyHtml = true;

            smtp.Port = 25;
            smtp.Host = "10.0.5.104";
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.EnableSsl = false;

            smtp.Send(message1);
        }


        private void BuildSKUERTable()
        {
            dtResult = new DataTable();
            dtResult.Columns.Add("Date");
            dtResult.Columns.Add("Customer");
            dtResult.Columns.Add("SKU Count");
            dtResult.Columns.Add("Total Duration");
            dtResult.Columns.Add("Effective Rate");
        }


        private void NonComplainceReportAsentinel()
        {
            DataTable dt = new System.Data.DataTable();
            //da = new SqlDataAdapter("select UL_Employee_ID as [Employee Id], R_User_Name as [Employee Name], T_TeamName as [Team Name], UL_RepMgrEmail as [Reporting Manager email Id],  max(R_SubmittedOn) as [Last Submitted Date] from RTM_Records, RTM_User_List, RTM_Team_List where R_User_Name = UL_User_Name and UL_Team_Id = T_ID and R_Submit =1 and R_SubmittedOn is not null and (UL_Hourly = 1 or UL_Hourly is null) and UL_Job_Title Not in ('Dir, Imp Shared Services', 'Dir, Program Management', 'Dir, Service Delivery', 'Dir, Strategic Consulting', 'Director, Audit', 'Director, CA & CB', 'Director, Content Management', 'Director, Engagement', 'Director, Financial Ops', 'Director, Implementations', 'Director, Operations', 'EVP, Operations', 'Sr Dir, Imp Onboarding', 'Vice President, Engagement', 'VP, Service Delivery', 'VP, Shared Services') and ([UL_Exclude] is null or [UL_Exclude] = 0) group by R_User_Name, UL_Employee_ID,T_TeamName, UL_RepMgrEmail order by T_TeamName,R_User_Name", con);
            string startDate1;
            startDate1 = DateTime.Now.AddDays(-10).ToString("MM/dd/yyyy");

            da = new SqlDataAdapter("select UL_Employee_ID as [Employee Id], R_User_Name as [Employee Name], T_TeamName as [Team Name], UL_RepMgrEmail as [Reporting Manager email Id], CONVERT(CHAR(10), max(R_Start_Date_Time), 101) as [Date Submitted Upto] from RTM_Records, RTM_User_List, RTM_Team_List where R_User_Name = UL_User_Name and UL_Team_Id = T_ID and R_Submit = '1' and (UL_Hourly = '0' or UL_Hourly is null) and UL_Job_Title Not in ('Dir, Imp Shared Services', 'Dir, Program Management', 'Dir, Service Delivery', 'Dir, Strategic Consulting', 'Director, Audit', 'Director, CA & CB', 'Director, Content Management', 'Director, Engagement', 'Director, Financial Ops', 'Director, Implementations', 'Director, Operations', 'EVP, Operations', 'Sr Dir, Imp Onboarding', 'Vice President, Engagement', 'VP, Service Delivery', 'VP, Shared Services') and ([UL_Exclude] is null or [UL_Exclude] = '0') and (UL_DOJ <= '" + startDate1 + "' or UL_DOJ is null) and UL_User_Status = '1' and T_Location = 'Asentinel' group by R_User_Name, UL_Employee_ID,T_TeamName, UL_RepMgrEmail order by R_User_Name,T_TeamName", con);

            da.Fill(dt);

            DataTable dtResult1 = new DataTable();
            string startDate;
            startDate = DateTime.Now.AddDays(-10).ToString("MM/dd/yyyy");
            //startDate = DateTime.Now.AddDays(-10);
            DataTable dtCloned = dt.Clone();
            dtCloned.Columns["Date Submitted Upto"].DataType = typeof(DateTime);
            foreach (DataRow row in dt.Rows)
            {
                dtCloned.ImportRow(row);
            }
            //dtResult.Columns["Date Submitted Upto"].DataType = typeof(DateTime);
            dtResult1 = dtCloned.Select("[Date Submitted Upto] <= '" + Convert.ToDateTime(startDate) + "'").CopyToDataTable();


            try
            {
                sendmailToRTMTeam("Mohammed.Shiddique@Tangoe.com", "NonComplianceReport Started");

                sendmailToDirectorsAsentinel("karen.donnelly@tangoe.com", dtResult1);

                sendmailToRTMTeam("Mohammed.Shiddique@Tangoe.com", "NonComplianceReport Completed");
            }

            catch (Exception ex)
            {
                sendmailToRTMTeam("Mohammed.Shiddique@Tangoe.com", ex.ToString());
            }
        }


        private void NonComplainceReport()
        {
            DataTable dt = new System.Data.DataTable();
            //da = new SqlDataAdapter("select UL_Employee_ID as [Employee Id], R_User_Name as [Employee Name], T_TeamName as [Team Name], UL_RepMgrEmail as [Reporting Manager email Id],  max(R_SubmittedOn) as [Last Submitted Date] from RTM_Records, RTM_User_List, RTM_Team_List where R_User_Name = UL_User_Name and UL_Team_Id = T_ID and R_Submit =1 and R_SubmittedOn is not null and (UL_Hourly = 1 or UL_Hourly is null) and UL_Job_Title Not in ('Dir, Imp Shared Services', 'Dir, Program Management', 'Dir, Service Delivery', 'Dir, Strategic Consulting', 'Director, Audit', 'Director, CA & CB', 'Director, Content Management', 'Director, Engagement', 'Director, Financial Ops', 'Director, Implementations', 'Director, Operations', 'EVP, Operations', 'Sr Dir, Imp Onboarding', 'Vice President, Engagement', 'VP, Service Delivery', 'VP, Shared Services') and ([UL_Exclude] is null or [UL_Exclude] = 0) group by R_User_Name, UL_Employee_ID,T_TeamName, UL_RepMgrEmail order by T_TeamName,R_User_Name", con);
            string startDate1;
            startDate1 = DateTime.Now.AddDays(-10).ToString("MM/dd/yyyy");

            da = new SqlDataAdapter("select UL_Employee_ID as [Employee Id], R_User_Name as [Employee Name], T_TeamName as [Team Name], UL_RepMgrEmail as [Reporting Manager email Id], CONVERT(CHAR(10), max(R_Start_Date_Time), 101) as [Date Submitted Upto] from RTM_Records, RTM_User_List, RTM_Team_List where R_User_Name = UL_User_Name and UL_Team_Id = T_ID and R_Submit = '1' and (UL_Hourly = '0' or UL_Hourly is null) and UL_Job_Title Not in ('Dir, Imp Shared Services', 'Dir, Program Management', 'Dir, Service Delivery', 'Dir, Strategic Consulting', 'Director, Audit', 'Director, CA & CB', 'Director, Content Management', 'Director, Engagement', 'Director, Financial Ops', 'Director, Implementations', 'Director, Operations', 'EVP, Operations', 'Sr Dir, Imp Onboarding', 'Vice President, Engagement', 'VP, Service Delivery', 'VP, Shared Services') and ([UL_Exclude] is null or [UL_Exclude] = '0') and (UL_DOJ <= '" + startDate1 + "' or UL_DOJ is null) and UL_User_Status = '1' group by R_User_Name, UL_Employee_ID,T_TeamName, UL_RepMgrEmail order by R_User_Name,T_TeamName", con);

            da.Fill(dt);

            DataRow dr;
            DataTable dtDirector = new DataTable();

            dtResult = new DataTable();
            dtResult.Columns.Add("Employee Id");
            dtResult.Columns.Add("Employee Name");
            dtResult.Columns.Add("Team Name");
            dtResult.Columns.Add("Reporting Manager email Id");
            dtResult.Columns.Add("Date Submitted Upto");
            dtResult.Columns.Add("Director");

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow drRow in dt.Rows)
                {
                    dr = dtResult.NewRow();

                    dr["Employee Id"] = drRow["Employee Id"];
                    dr["Employee Name"] = drRow["Employee Name"];
                    dr["Team Name"] = drRow["Team Name"];
                    dr["Reporting Manager email Id"] = drRow["Reporting Manager email Id"];
                    dr["Date Submitted Upto"] = drRow["Date Submitted Upto"];
                    try
                    {
                        dtDirector = getDirectorFromMasterTable(drRow["Employee Id"].ToString());
                    }
                    catch
                    {

                    }

                    //dtDirector = objTSheet.getDirector(drRow["Employee Name"].ToString());
                    if (dtDirector.Rows.Count > 0)
                    {
                        string expression = "MUL_Job_Title Like '%Dir%'";
                        string sortOrder = "Lvl ASC";
                        DataRow[] foundRows;
                        foundRows = dtDirector.Select(expression, sortOrder);
                        if (foundRows.Length > 0)
                        {
                            dr["Director"] = foundRows[0]["MUL_EmailId"];
                        }
                        else
                        {
                            expression = "MUL_Job_Title Like '%VP%'";
                            foundRows = dtDirector.Select(expression, sortOrder);
                            if (foundRows.Length > 0)
                            {
                                dr["Director"] = foundRows[0]["MUL_EmailId"]; // dtDirector.Rows[0]["UL_EmailId"].ToString();
                            }
                            else
                            {
                                dr["Director"] = "";
                            }
                        }
                    }
                    else
                    {
                        dr["Director"] = "";
                    }

                    dtResult.Rows.Add(dr);
                }

                DataTable dtResult1 = new DataTable();
                string startDate;
                startDate = DateTime.Now.AddDays(-10).ToString("MM/dd/yyyy");
                //startDate = DateTime.Now.AddDays(-10);
                DataTable dtCloned = dtResult.Clone();
                dtCloned.Columns["Date Submitted Upto"].DataType = typeof(DateTime);
                foreach (DataRow row in dtResult.Rows)
                {
                    dtCloned.ImportRow(row);
                }
                //dtResult.Columns["Date Submitted Upto"].DataType = typeof(DateTime);
                dtResult1 = dtCloned.Select("[Date Submitted Upto] <= '" + Convert.ToDateTime(startDate) + "'").CopyToDataTable();

                //string filePath = "" + Directory.GetCurrentDirectory() + "\\nonCompliance.csv";
                //CSVUtility.ToCSV(dtResult1, filePath);
                DataTable dtUnique = dtResult1.DefaultView.ToTable(true, "Director");

                //dtResult1.Rows.Add(result);

                try
                {
                    sendmailToRTMTeam("Mohammed.Shiddique@Tangoe.com", "NonComplianceReport Started");
                    foreach (DataRow drRow in dtUnique.Rows)
                    {
                        if (drRow["Director"].ToString() != "")
                        {
                            //if (drRow["Director"].ToString().Replace("@tangoe.com", "").Replace(".", " ") == "carol villa" | drRow["Director"].ToString().Replace("@tangoe.com", "").Replace(".", " ") == "james jones" | drRow["Director"].ToString().Replace("@tangoe.com", "").Replace(".", " ") == "amy densmore")
                            //{


                            DataTable dtResult2 = dtResult1.Select("Director = '" + drRow["Director"].ToString() + "'", "[TEAM NAME],[Date Submitted Upto] ASC").CopyToDataTable();

                            dtResult2.Columns.RemoveAt(0);
                            dtResult2.Columns.RemoveAt(2);
                            dtResult2.Columns.RemoveAt(3);

                            DataTable dtResult3 = new DataTable();
                            dtResult3.Columns.Add("S. NO");
                            dtResult3.Columns.Add("EMPLOYEE NAME");
                            dtResult3.Columns.Add("TEAM NAME");
                            dtResult3.Columns.Add("DATE SUBMITTED UPTO");
                            dtResult3.Columns.Add("NO. OF TIMES SHEET SUBMISSION MISSING");
                            int i = 0;
                            foreach (DataRow drRow1 in dtResult2.Rows)
                            {
                                i += 1;

                                int nosOfWeek = 0;
                                try
                                {
                                    if (DateTime.Now > Convert.ToDateTime(drRow1["DATE SUBMITTED UPTO"]))
                                    {
                                        nosOfWeek = NumberOfWeeks(Convert.ToDateTime(drRow1["DATE SUBMITTED UPTO"]).Date, DateTime.Today);
                                    }
                                }
                                catch { }

                                dtResult3.Rows.Add(i, drRow1["EMPLOYEE NAME"], drRow1["TEAM NAME"], Convert.ToDateTime(drRow1["DATE SUBMITTED UPTO"].ToString()).ToShortDateString(), nosOfWeek);
                            }

                            //dtResult2.Columns.RemoveAt(4);

                            sendmailToDirectors(drRow["Director"].ToString(), dtResult3);
                        }
                    }

                    sendmailToRTMTeam("Mohammed.Shiddique@Tangoe.com", "NonComplianceReport Completed");
                }

                catch (Exception ex)
                {
                    sendmailToRTMTeam("Mohammed.Shiddique@Tangoe.com", ex.ToString());
                }


            }


        }

        public static int NumberOfWeeks(DateTime dateFrom, DateTime dateTo)
        {
            TimeSpan Span = dateTo.Subtract(dateFrom);
            int Days = Span.Days - 7 + (int)dateFrom.DayOfWeek;
            int WeekCount = Days / 7;
            return WeekCount;
        }


        public DataTable getDirectorFromMasterTable(string empId)
        {
            string sQuery;
            dt = new DataTable();
            //SqlParameter[] param = new SqlParameter[]{
            //    new SqlParameter("@empId", empId)
            //};

            sQuery = ";WITH CTE_Traverse_hierarchy " +
                        "AS " +
                        "( " +
                          " SELECT MUL_EmployeeId,MUL_EmailId,MUL_ManagerId,MUL_Job_Title,MUL_ManagerEmail_Id,0 Lvl FROM RTM_Master_UserList Where MUL_EmailId != MUL_ManagerEmail_Id and MUL_EmployeeId=@empId " +
                           "UNION ALL " +
                           "SELECT E.MUL_EmployeeId,E.MUL_EmailId,E.MUL_ManagerId,E.MUL_Job_Title,E.MUL_ManagerEmail_Id,Lvl+1 Lvl FROM RTM_Master_UserList E " +
                           "JOIN CTE_Traverse_hierarchy Parent on Parent.MUL_ManagerId=E.MUL_EmployeeId WHERE E.MUL_EmailId != E.MUL_ManagerEmail_Id " +
                        ") " +
                        "Select * from CTE_Traverse_hierarchy";

            da = new SqlDataAdapter(sQuery, con);

            da.SelectCommand.Parameters.Add(new SqlParameter
            {
                ParameterName = "@empId",
                Value = empId
            });
            da.Fill(dt);
            return dt;
        }


        private void sendmailToDirectors(string emailId, DataTable dt)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<p>Hi,</p>");
            sb.AppendLine("");
            sb.AppendLine("<p>Please find below employees who did not submit their RTM time sheet last week.</p>");
            sb.AppendLine("");
            sb.AppendLine(DataTableToHTMLTable(dt));
            //sb.AppendLine("<p>Day count : " + dayCount + ". Hours recorded : " + hours + "</p>");
            //sb.AppendLine(myBuilder.ToString());
            sb.AppendLine("");
            sb.AppendLine("<p>Regards,</p>");
            //sb.AppendLine("");
            sb.AppendLine("<p><b>RTM Support</b></p>");
            sb.AppendLine("");
            sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
            sb.AppendLine("");

            MailMessage message1 = new MailMessage();
            SmtpClient smtp = new SmtpClient();

            message1.From = new MailAddress(FromAddress);


            message1.Bcc.Add(new MailAddress("Sleema.Joseph@tangoe.com"));
            message1.Bcc.Add(new MailAddress("Mohammed.Shiddique@tangoe.com"));
            message1.Bcc.Add(new MailAddress("Vidhya.Karunagaran@tangoe.com"));
            //message1.Bcc.Add(new MailAddress("Hariram.R@tangoe.com"));

            //-------------Test code--------------
            //message1.To.Add(new MailAddress("Vidhya.Karunagaran@tangoe.com"));

            //-------------Production code--------------
            if (emailId == "jim.carroll@tangoe.com")
            {
                message1.To.Add(new MailAddress("diane.miller@tangoe.com"));
            }
            else
            {
                message1.To.Add(new MailAddress(emailId));
            }


            if (emailId == "amy.densmore@tangoe.com" || emailId == "Ben.Crees@eu.tangoe.com" || emailId == "BetsyR@tangoe.com" || emailId == "carol.villa@tangoe.com" || emailId == "Christopher.Dubanowitz@tangoe.com" || emailId == "deborah.hughes@tangoe.com" || emailId == "grete.mortley@tangoe.com" || emailId == "james.jones@tangoe.com" || emailId == "johagan@tangoe.com" || emailId == "Lauryn.Robinson@tangoe.com" || emailId == "mike.mirakian@tangoe.com" || emailId == "Patricia.Purtill@tangoe.com" || emailId == "renee.newton@tangoe.com" || emailId == "Shellie.Allen@tangoe.com" || emailId == "victoria.litchfield@eu.tangoe.com")
            {
                message1.CC.Add(new MailAddress("diane.miller@tangoe.com"));
            }
            else if (emailId == "Rashmi.Ahuja@tangoe.com")
            {
                message1.CC.Add(new MailAddress("AhujaDirects@tangoe.com"));
                message1.CC.Add(new MailAddress("Shabeenaz1@tangoe.com"));
            }

            message1.Subject = "Non Compliance Report - Last Week";

            message1.Body = sb.ToString();

            message1.IsBodyHtml = true;


            //SmtpClient smtpClient = new SmtpClient("mail.north.tangoe.com");
            //smtpClient.UseDefaultCredentials = false;
            //NetworkCredential credentials = new NetworkCredential("Mohammed.Shiddique", "Yunoosaug123!");
            //smtpClient.Credentials = credentials;
            //smtpClient.Send(message1);


            smtp.Port = 25;
            smtp.Host = "10.0.5.104";
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.EnableSsl = false;

            smtp.Send(message1);
        }

        private void sendmailToDirectorsAsentinel(string emailId, DataTable dt)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<p>Hi,</p>");
            sb.AppendLine("");
            sb.AppendLine("<p>Please find below employees who did not submit their RTM time sheet last week.</p>");
            sb.AppendLine("");
            sb.AppendLine(DataTableToHTMLTable(dt));
            //sb.AppendLine("<p>Day count : " + dayCount + ". Hours recorded : " + hours + "</p>");
            //sb.AppendLine(myBuilder.ToString());
            sb.AppendLine("");
            sb.AppendLine("<p>Regards,</p>");
            //sb.AppendLine("");
            sb.AppendLine("<p><b>RTM Support</b></p>");
            sb.AppendLine("");
            sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
            sb.AppendLine("");

            MailMessage message1 = new MailMessage();
            SmtpClient smtp = new SmtpClient();

            message1.From = new MailAddress(FromAddress);

            message1.To.Add(new MailAddress(emailId));


            //message1.Bcc.Add(new MailAddress("Sleema.Joseph@tangoe.com"));
            //message1.Bcc.Add(new MailAddress("Mohammed.Shiddique@tangoe.com"));
            message1.Bcc.Add(new MailAddress("RTM-Support@tangoe.com"));
            message1.CC.Add(new MailAddress("BetsyR@tangoe.com"));
            message1.CC.Add(new MailAddress("victoria.litchfield@eu.tangoe.com"));
            message1.CC.Add(new MailAddress("grete.mortley@tangoe.com"));
            message1.CC.Add(new MailAddress("diane.miller@tangoe.com"));


            message1.Subject = "Non Compliance Report";

            message1.Body = sb.ToString();

            message1.IsBodyHtml = true;
            smtp.Port = 25;
            smtp.Host = "10.0.5.104";
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.EnableSsl = false;

            smtp.Send(message1);
        }

        public string DataTableToHTMLTable(DataTable inTable)
        {
            StringBuilder dString = new StringBuilder();
            dString.Append("<table border=1 cellpadding=1 cellspacing=0 style= border: solid 1px Silver;font-size:15pt;font-family:sans-serif;> ");
            dString.Append(GetHeader(inTable));
            dString.Append(GetBody(inTable));
            //dString.Append("</tbody>")
            dString.Append("</table>");
            return dString.ToString();
        }

        private string GetHeader(DataTable dTable)
        {
            StringBuilder dString = new StringBuilder();

            //dString.Append("<tr border='1px' ")
            //dString.Append("style='border: solid 1px Black; font-size: small;'>")
            dString.Append("<tr style=font-variant:small-caps;font-style:Bold;color:Black;font-size:15px;>");
            foreach (DataColumn dColumn in dTable.Columns)
            {
                dString.Append("<th bgcolor=#FFFF00>");
                dString.AppendFormat(dColumn.ColumnName);
                dString.Append("</th>");
            }
            dString.Append("</tr>");

            return dString.ToString();
        }

        private string GetBody(DataTable dTable)
        {

            //try
            //{
            StringBuilder dString = new StringBuilder();

            foreach (DataRow dRow in dTable.Rows)
            {
                dString.Append("<tr style=font-variant:small-caps;font-style:normal;color:Black;font-size:15px;>");
                for (int dCount = 0; dCount <= dTable.Columns.Count - 1; dCount++)
                {
                    dString.Append("<td bgcolor=#FFFFFF>");
                    dString.AppendFormat(dRow[dCount].ToString());
                    dString.Append("</td>");
                }
                dString.Append("</tr>");
            }


            return dString.ToString();


            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());

            //}
        }

        private void SendSKUEffectiveRate()
        {
            DataRow dr;
            BuildSKUERTable();
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            dt = new DataTable();
            string sQuery = "";
            if (indianTime.DayOfWeek == DayOfWeek.Monday)
            {
                sQuery = "select Count(SKU_ID) as [SKU], CL_ClientName as Client,SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),1) as float))/3600) as [Total Duration]  from RTM_IPVDetails, RTM_Records, RTM_Client_List where RTM_IPVDetails.R_Id = RTM_Records.R_ID and R_Client = CL_ID and (SubTask_Id='129' or SubTask_Id='141' or SubTask_Id ='1740') and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, EndTime))) ='" + indianTime.AddDays(-3).ToShortDateString() + "'  GROUP BY CL_ClientName";
            }
            else
            {
                sQuery = "select Count(SKU_ID) as [SKU], CL_ClientName as Client,SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),1) as float))/3600) as [Total Duration]  from RTM_IPVDetails, RTM_Records, RTM_Client_List where RTM_IPVDetails.R_Id = RTM_Records.R_ID and R_Client = CL_ID and (SubTask_Id='129' or SubTask_Id='141' or SubTask_Id ='1740') and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, EndTime))) ='" + indianTime.AddDays(-1).ToShortDateString() + "'  GROUP BY CL_ClientName";
            }

            using (da = new SqlDataAdapter(sQuery, con))
            {
                da.Fill(dt);
            }
            double totalDuration = 0;
            int skuCount = 0;
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow drRow in dt.Rows)
                {
                    dr = dtResult.NewRow();

                    dr["Date"] = indianTime.AddDays(-1).ToShortDateString();
                    dr["Customer"] = drRow["Client"];
                    dr["SKU Count"] = drRow["SKU"];
                    skuCount = skuCount + Convert.ToInt32(drRow["SKU"]);
                    double totalDur = Math.Round(Convert.ToDouble(drRow["Total Duration"]), 2, MidpointRounding.AwayFromZero);
                    totalDuration = totalDuration + totalDur;
                    dr["Total Duration"] = totalDur;
                    if (totalDur > 0)
                    {
                        dr["Effective Rate"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SKU"]) / totalDur), 2, MidpointRounding.AwayFromZero);
                    }
                    else
                    {
                        dr["Effective Rate"] = "";
                    }

                    dtResult.Rows.Add(dr);

                }

                dr = dtResult.NewRow();

                dr["Date"] = "";
                dr["Customer"] = "Total";
                dr["SKU Count"] = skuCount;
                dr["Total Duration"] = totalDuration;
                dr["Effective Rate"] = Math.Round(skuCount / totalDuration, 2, MidpointRounding.AwayFromZero);

                dtResult.Rows.Add(dr);
            }

            if (dtResult.Rows.Count > 0)
            {
                getDelayHTML(dtResult);
                MailMessage message1 = new MailMessage();
                //string filePath = "" + Directory.GetCurrentDirectory() + "\\Effective Rate Invoices - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv";
                //CSVUtility.ToCSV(dtResult, filePath);

                DirectCSV csv = new DirectCSV();
                var data = csv.ExportToCSV(dtResult);
                var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(data);
                MemoryStream ms = new MemoryStream(bytes);

                Attachment attachFile = new Attachment(ms, "Effective Rate Invoices - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv", "application/csv");
                message1.Attachments.Add(attachFile);

                StringBuilder sb = new StringBuilder();

                sb.AppendLine("");
                sb.AppendLine(myBuilder.ToString());
                sb.AppendLine("");//here I want the data to       display in table format
                sb.AppendLine("");
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                sb.AppendLine("");


                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);

                message1.To.Add(new MailAddress("Sriram.Krishnan@tangoe.com"));
                message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                message1.To.Add(new MailAddress("Sandesh.Ravichandra@tangoe.com"));
                message1.To.Add(new MailAddress("Johwessly.Chennaiah@tangoe.com"));
                message1.CC.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));

                message1.Subject = "Effective Rate Invoices - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy");

                message1.Body = sb.ToString();
                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment(filePath);
                //message1.Attachments.Add(attachment);
                message1.IsBodyHtml = true;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                //smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }


        }

        private void BuildSKUDBERTable()
        {
            dtResult = new DataTable();
            dtResult.Columns.Add("Date");
            dtResult.Columns.Add("Client Name");
            dtResult.Columns.Add("No of Invoices");
            dtResult.Columns.Add("IP Time");
            dtResult.Columns.Add("QC Time");
            dtResult.Columns.Add("Total Time");
            dtResult.Columns.Add("Effective Rate - IP");
            dtResult.Columns.Add("Effective Rate - QC");
            dtResult.Columns.Add("Effective Rate - Overall");
        }

        //Effective Rate Invoices - IPV 
        private void SendSKUEffectiveRateFromSKUDB()
        {
            try
            {
                DataRow dr;
                BuildSKUDBERTable();
                DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
                dt = new DataTable();
                string sQuery = "";

                sQuery = "select COUNT(SKU_NUMBER) as [SKU], CL_CLientName as Client from RTM_Sku " +
                            "left join RTM_Client_List on TSHEETS_CLIENT_CODE = CL_Code " +
                            "left join RTM_User_List on fullname = RIGHT(UL_System_User_Name, LEN(UL_System_User_Name) - 5) " +
                            "WHERE CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, DATE_FINISHED))) ='" + indianTime.AddDays(-1).ToShortDateString() + "' and CL_TeamId = 29 and UL_Team_ID = 29 and CL_Status = 1 and CL_Product = 'IPV' " +
                            "Group by CL_ClientName";

                using (da = new SqlDataAdapter(sQuery, con))
                {
                    da.Fill(dt);
                }
                double totalIPDuration = 0;
                double totalQCDuration = 0;
                int skuCount = 0;
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow drRow in dt.Rows)
                    {
                        dr = dtResult.NewRow();
                        dr["Date"] = indianTime.AddDays(-1).ToShortDateString();
                        dr["Client Name"] = drRow["Client"];
                        dr["No of Invoices"] = drRow["SKU"];
                        skuCount = skuCount + Convert.ToInt32(drRow["SKU"]);
                        DataTable dtDuration = new DataTable();
                        dtDuration = getSKUDuration(drRow["Client"].ToString(), indianTime.AddDays(-1).ToShortDateString());
                        double totalIPDur = 0;
                        double totalQCDur = 0;
                        if (dtDuration.Rows.Count > 0)
                        {
                            totalIPDur = Math.Round(Convert.ToDouble(dtDuration.Rows[0]["IPtime"]), 2, MidpointRounding.AwayFromZero);
                            totalIPDuration = totalIPDuration + totalIPDur;

                            totalQCDur = Math.Round(Convert.ToDouble(dtDuration.Rows[0]["QCtime"]), 2, MidpointRounding.AwayFromZero);
                            totalQCDuration = totalQCDuration + totalQCDur;
                        }
                        else
                        {
                            totalIPDur = 0;
                            totalQCDur = 0;
                        }

                        dr["IP Time"] = totalIPDur;
                        dr["QC Time"] = totalQCDur;
                        double totalTime = totalIPDur + totalQCDur;
                        dr["Total Time"] = Math.Round((totalIPDur + totalQCDur), 2, MidpointRounding.AwayFromZero);
                        if (totalIPDur > 0)
                        {
                            dr["Effective Rate - IP"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SKU"]) / totalIPDur), 2, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            dr["Effective Rate - IP"] = "";
                        }

                        if (totalQCDur > 0)
                        {
                            dr["Effective Rate - QC"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SKU"]) / totalQCDur), 2, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            dr["Effective Rate - QC"] = "";
                        }

                        if (totalTime > 0)
                        {
                            dr["Effective Rate - Overall"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SKU"]) / totalTime), 2, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            dr["Effective Rate - Overall"] = "";
                        }

                        dtResult.Rows.Add(dr);
                    }

                    dr = dtResult.NewRow();

                    dr["Date"] = "";
                    dr["Client Name"] = "Total";
                    dr["No of Invoices"] = skuCount;
                    dr["IP Time"] = totalIPDuration;
                    dr["QC Time"] = totalQCDuration;
                    dr["Total Time"] = Math.Round((totalIPDuration + totalQCDuration), 2, MidpointRounding.AwayFromZero);
                    dr["Effective Rate - IP"] = Math.Round(skuCount / totalIPDuration, 2, MidpointRounding.AwayFromZero);
                    dr["Effective Rate - QC"] = Math.Round(skuCount / totalQCDuration, 2, MidpointRounding.AwayFromZero);
                    double grandTotal = Math.Round((totalIPDuration + totalQCDuration), 2, MidpointRounding.AwayFromZero);
                    dr["Effective Rate - Overall"] = Math.Round(Convert.ToDouble(skuCount / grandTotal), 2, MidpointRounding.AwayFromZero);
                    dtResult.Rows.Add(dr);
                }

                if (dtResult.Rows.Count > 0)
                {
                    getDelayHTML(dtResult);

                    //string filePath = "" + Directory.GetCurrentDirectory() + "\\Effective Rate Invoices - IPV - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv";
                    //CSVUtility.ToCSV(dtResult, filePath);

                    StringBuilder sb = new StringBuilder();

                    sb.AppendLine("");
                    sb.AppendLine(myBuilder.ToString());
                    sb.AppendLine("");//here I want the data to       display in table format
                    sb.AppendLine("");
                    sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                    sb.AppendLine("");

                    MailMessage message1 = new MailMessage();

                    DirectCSV csv = new DirectCSV();
                    var data = csv.ExportToCSV(dtResult);
                    var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(data);
                    MemoryStream ms = new MemoryStream(bytes);

                    Attachment attachFile = new Attachment(ms, "Effective Rate Invoices - IPV - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv", "application/csv");
                    message1.Attachments.Add(attachFile);

                    SmtpClient smtp = new SmtpClient();

                    message1.From = new MailAddress(FromAddress);

                    message1.To.Add(new MailAddress("Sriram.Krishnan@tangoe.com"));
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Sandesh.Ravichandra@tangoe.com"));
                    message1.To.Add(new MailAddress("Johwessly.Chennaiah@tangoe.com"));
                    message1.CC.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    //message1.CC.Add(new MailAddress("rich.lena@tangoe.com"));
                    //message1.CC.Add(new MailAddress("melissa.guarracino@tangoe.com"));
                    message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));

                    message1.Subject = "Effective Rate Invoices - IPV - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy");

                    message1.Body = sb.ToString();
                    //System.Net.Mail.Attachment attachment;
                    //attachment = new System.Net.Mail.Attachment(filePath);
                    //message1.Attachments.Add(attachment);
                    message1.IsBodyHtml = true;

                    smtp.Port = 25;
                    smtp.Host = "10.0.5.104";
                    //smtp.Host = "outlook-south.tangoe.com";
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    smtp.EnableSsl = false;
                    smtp.Send(message1);
                }
            }
            catch (Exception)
            {


            }

        }
        //Effective Rate Invoices - IPV -working
        private void SendSampleSKUEffectiveRateFromSKUDB()
        {
            try
            {
                DataRow dr;
                BuildSKUDBERTable();
                DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
                dt = new DataTable();
                string sQuery = "";

                sQuery = "select COUNT(SKU_NUMBER) as [SKU], CL_CLientName as Client from RTM_Sku " +
                            "left join RTM_Client_List on TSHEETS_CLIENT_CODE = CL_Code " +
                            "left join RTM_User_List on fullname = RIGHT(UL_System_User_Name, LEN(UL_System_User_Name) - 5) " +
                            "WHERE CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, DATE_FINISHED))) ='" + indianTime.AddDays(-1).ToShortDateString() + "' and CL_TeamId = 29 and UL_Team_ID = 29 and CL_Status = 1 and CL_Product = 'IPV' " +
                            "Group by CL_ClientName";

                using (da = new SqlDataAdapter(sQuery, con))
                {
                    da.Fill(dt);
                }
                double totalIPDuration = 0;
                double totalQCDuration = 0;
                int skuCount = 0;
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow drRow in dt.Rows)
                    {
                        dr = dtResult.NewRow();
                        dr["Date"] = indianTime.AddDays(-1).ToShortDateString();
                        dr["Client Name"] = drRow["Client"];
                        dr["No of Invoices"] = drRow["SKU"];
                        skuCount = skuCount + Convert.ToInt32(drRow["SKU"]);
                        DataTable dtDuration = new DataTable();
                        dtDuration = getSKUDuration(drRow["Client"].ToString(), indianTime.AddDays(-1).ToShortDateString());
                        double totalIPDur = 0;
                        double totalQCDur = 0;
                        if (dtDuration.Rows.Count > 0)
                        {
                            totalIPDur = Math.Round(Convert.ToDouble(dtDuration.Rows[0]["IPtime"]), 2, MidpointRounding.AwayFromZero);
                            totalIPDuration = totalIPDuration + totalIPDur;

                            totalQCDur = Math.Round(Convert.ToDouble(dtDuration.Rows[0]["QCtime"]), 2, MidpointRounding.AwayFromZero);
                            totalQCDuration = totalQCDuration + totalQCDur;
                        }
                        else
                        {
                            totalIPDur = 0;
                            totalQCDur = 0;
                        }

                        dr["IP Time"] = totalIPDur;
                        dr["QC Time"] = totalQCDur;
                        double totalTime = totalIPDur + totalQCDur;
                        dr["Total Time"] = Math.Round((totalIPDur + totalQCDur), 2, MidpointRounding.AwayFromZero);
                        if (totalIPDur > 0)
                        {
                            dr["Effective Rate - IP"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SKU"]) / totalIPDur), 2, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            dr["Effective Rate - IP"] = "";
                        }

                        if (totalQCDur > 0)
                        {
                            dr["Effective Rate - QC"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SKU"]) / totalQCDur), 2, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            dr["Effective Rate - QC"] = "";
                        }

                        if (totalTime > 0)
                        {
                            dr["Effective Rate - Overall"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SKU"]) / totalTime), 2, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            dr["Effective Rate - Overall"] = "";
                        }

                        dtResult.Rows.Add(dr);
                    }

                    dr = dtResult.NewRow();

                    dr["Date"] = "";
                    dr["Client Name"] = "Total";
                    dr["No of Invoices"] = skuCount;
                    dr["IP Time"] = totalIPDuration;
                    dr["QC Time"] = totalQCDuration;
                    dr["Total Time"] = Math.Round((totalIPDuration + totalQCDuration), 2, MidpointRounding.AwayFromZero);
                    dr["Effective Rate - IP"] = Math.Round(skuCount / totalIPDuration, 2, MidpointRounding.AwayFromZero);
                    dr["Effective Rate - QC"] = Math.Round(skuCount / totalQCDuration, 2, MidpointRounding.AwayFromZero);
                    double grandTotal = Math.Round((totalIPDuration + totalQCDuration), 2, MidpointRounding.AwayFromZero);
                    dr["Effective Rate - Overall"] = Math.Round(Convert.ToDouble(skuCount / grandTotal), 2, MidpointRounding.AwayFromZero);
                    dtResult.Rows.Add(dr);
                }

                if (dtResult.Rows.Count > 0)
                {
                    getDelayHTML(dtResult);

                    //string filePath = "" + Directory.GetCurrentDirectory() + "\\Effective Rate Invoices - IPV - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv";
                    //CSVUtility.ToCSV(dtResult, filePath);

                    StringBuilder sb = new StringBuilder();

                    sb.AppendLine("");
                    sb.AppendLine(myBuilder.ToString());
                    sb.AppendLine("");//here I want the data to       display in table format
                    sb.AppendLine("");
                    sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                    sb.AppendLine("");

                    MailMessage message1 = new MailMessage();

                    DirectCSV csv = new DirectCSV();
                    var data = csv.ExportToCSV(dtResult);
                    var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(data);
                    MemoryStream ms = new MemoryStream(bytes);

                    Attachment attachFile = new Attachment(ms, "Effective Rate Invoices - IPV - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv", "application/csv");
                    message1.Attachments.Add(attachFile);

                    SmtpClient smtp = new SmtpClient();

                    message1.From = new MailAddress(FromAddress);

                    message1.To.Add(new MailAddress("Sriram.Krishnan@tangoe.com"));
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Sandesh.Ravichandra@tangoe.com"));
                    message1.To.Add(new MailAddress("Johwessly.Chennaiah@tangoe.com"));
                    message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));

                    message1.Subject = "Effective Rate Invoices - IPV - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy");

                    message1.Body = sb.ToString();
                    //System.Net.Mail.Attachment attachment;
                    //attachment = new System.Net.Mail.Attachment(filePath);
                    //message1.Attachments.Add(attachment);
                    message1.IsBodyHtml = true;

                    smtp.Port = 25;
                    smtp.Host = "10.0.5.104";
                    //smtp.Host = "outlook-south.tangoe.com";
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    smtp.EnableSsl = false;

                    smtp.Send(message1);
                }
            }
            catch (Exception)
            {


            }

        }

        private void BuildCMPDBERTable()
        {
            dtResult = new DataTable();
            dtResult.Columns.Add("Date");
            dtResult.Columns.Add("Client Name");
            dtResult.Columns.Add("No of Invoices");
            dtResult.Columns.Add("IP Time");
            dtResult.Columns.Add("QC Time");
            dtResult.Columns.Add("Total Time");
            dtResult.Columns.Add("Effective Rate - IP");
            dtResult.Columns.Add("Effective Rate - QC");
            dtResult.Columns.Add("Effective Rate - Overall");
        }

        private void SendCMPEffectiveRateFromCMPDB()
        {
            try
            {
                DataRow dr;
                BuildSKUDBERTable();
                DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
                dt = new DataTable();
                string sQuery = "";

                sQuery = "select COUNT(invoiceSubId) as [SID], custAbbr as [Client Code], CL_ClientName as [Client] from RTM_CMP " +
                        "left join RTM_Client_List on custAbbr = CL_Code " +
                        "where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, buildDateBackend))) = '" + indianTime.AddDays(-1).ToShortDateString() + "' and CL_TeamId = 29 and CL_Status = 1 and CL_Product ='CMP' " +
                        "group by custAbbr, CL_ClientName";

                //sQuery = "select COUNT(invoiceSubId) as [SID], custAbbr as [Client Code], CL_ClientName as [Client] from RTM_CMP " +
                //        "left join RTM_Client_List on custAbbr = CL_Code " +
                //        "where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, buildDateBackend))) = '09/20/2017' and CL_TeamId = 29 and CL_Status = 1 and CL_Product ='CMP' " +
                //        "group by custAbbr, CL_ClientName";

                using (da = new SqlDataAdapter(sQuery, con))
                {
                    da.Fill(dt);
                }
                double totalIPDuration = 0;
                double totalQCDuration = 0;
                int skuCount = 0;

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow drRow in dt.Rows)
                    {
                        dr = dtResult.NewRow();
                        dr["Date"] = indianTime.AddDays(-1).ToShortDateString();
                        dr["Client Name"] = drRow["Client"];
                        dr["No of Invoices"] = drRow["SID"];
                        skuCount = skuCount + Convert.ToInt32(drRow["SID"]);
                        DataTable dtDuration = new DataTable();
                        dtDuration = getCMPDuration(drRow["Client"].ToString(), indianTime.AddDays(-1).ToShortDateString());
                        //dtDuration = getCMPDuration(drRow["Client"].ToString(), "09/20/2017");
                        double totalIPDur = 0;
                        double totalQCDur = 0;
                        if (dtDuration.Rows.Count > 0)
                        {
                            totalIPDur = Math.Round(Convert.ToDouble(dtDuration.Rows[0]["IPtime"]), 2, MidpointRounding.AwayFromZero);
                            totalIPDuration = totalIPDuration + totalIPDur;

                            totalQCDur = Math.Round(Convert.ToDouble(dtDuration.Rows[0]["QCtime"]), 2, MidpointRounding.AwayFromZero);
                            totalQCDuration = totalQCDuration + totalQCDur;
                        }
                        else
                        {
                            totalIPDur = 0;
                            totalQCDur = 0;
                        }

                        dr["IP Time"] = totalIPDur;
                        dr["QC Time"] = totalQCDur;
                        double totalTime = totalIPDur + totalQCDur;
                        dr["Total Time"] = Math.Round((totalIPDur + totalQCDur), 2, MidpointRounding.AwayFromZero);
                        if (totalIPDur > 0)
                        {
                            dr["Effective Rate - IP"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SID"]) / totalIPDur), 2, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            dr["Effective Rate - IP"] = "";
                        }

                        if (totalQCDur > 0)
                        {
                            dr["Effective Rate - QC"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SID"]) / totalQCDur), 2, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            dr["Effective Rate - QC"] = "";
                        }

                        if (totalTime > 0)
                        {
                            dr["Effective Rate - Overall"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SID"]) / totalTime), 2, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            dr["Effective Rate - Overall"] = "";
                        }

                        dtResult.Rows.Add(dr);
                    }

                    dr = dtResult.NewRow();

                    dr["Date"] = "";
                    dr["Client Name"] = "Total";
                    dr["No of Invoices"] = skuCount;
                    dr["IP Time"] = totalIPDuration;
                    dr["QC Time"] = totalQCDuration;
                    dr["Total Time"] = Math.Round((totalIPDuration + totalQCDuration), 2, MidpointRounding.AwayFromZero);
                    dr["Effective Rate - IP"] = Math.Round(skuCount / totalIPDuration, 2, MidpointRounding.AwayFromZero);
                    dr["Effective Rate - QC"] = Math.Round(skuCount / totalQCDuration, 2, MidpointRounding.AwayFromZero);
                    double grandTotal = Math.Round((totalIPDuration + totalQCDuration), 2, MidpointRounding.AwayFromZero);
                    dr["Effective Rate - Overall"] = Math.Round(Convert.ToDouble(skuCount / grandTotal), 2, MidpointRounding.AwayFromZero);
                    dtResult.Rows.Add(dr);
                }

                if (dtResult.Rows.Count > 0)
                {
                    getDelayHTML(dtResult);

                    //string filePath = "" + Directory.GetCurrentDirectory() + "\\Effective Rate Invoices - CMP - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv";
                    //CSVUtility.ToCSV(dtResult, filePath);

                    StringBuilder sb = new StringBuilder();

                    sb.AppendLine("");
                    sb.AppendLine(myBuilder.ToString());
                    sb.AppendLine("");//here I want the data to       display in table format
                    sb.AppendLine("");
                    sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                    sb.AppendLine("");

                    MailMessage message1 = new MailMessage();

                    DirectCSV csv = new DirectCSV();
                    var data = csv.ExportToCSV(dtResult);
                    var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(data);
                    MemoryStream ms = new MemoryStream(bytes);

                    Attachment attachFile = new Attachment(ms, "Effective Rate Invoices - CMP - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv", "application/csv");
                    message1.Attachments.Add(attachFile);

                    SmtpClient smtp = new SmtpClient();

                    message1.From = new MailAddress(FromAddress);

                    message1.To.Add(new MailAddress("Sriram.Krishnan@tangoe.com"));
                    message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                    message1.To.Add(new MailAddress("Gurumanjesh.Gangadhara@tangoe.com"));
                    message1.To.Add(new MailAddress("Johwessly.Chennaiah@tangoe.com"));
                    message1.CC.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                    //message1.CC.Add(new MailAddress("rich.lena@tangoe.com"));
                    message1.CC.Add(new MailAddress("melissa.guarracino@tangoe.com"));
                    message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));

                    message1.Subject = "Effective Rate Invoices - CMP - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy");

                    message1.Body = sb.ToString();
                    //System.Net.Mail.Attachment attachment;
                    //attachment = new System.Net.Mail.Attachment(filePath);
                    //message1.Attachments.Add(attachment);
                    message1.IsBodyHtml = true;

                    smtp.Port = 25;
                    smtp.Host = "10.0.5.104";
                    //smtp.Host = "outlook-south.tangoe.com";
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    smtp.EnableSsl = false;

                    smtp.Send(message1);
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
        }

        private void BuildOrdersERTable()
        {
            dtResult = new DataTable();
            dtResult.Columns.Add("Date");
            dtResult.Columns.Add("Client");
            dtResult.Columns.Add("No of Orders");
            dtResult.Columns.Add("Total Time");
            dtResult.Columns.Add("Effective Rate");
        }

        private void SendOrdersEffectiveRate()
        {
            DataRow dr;
            BuildOrdersERTable();
            double totalCount = 0;
            double totaltime = 0;
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            dt = new DataTable();
            string sQuery = "";

            sQuery = "select	CAST(TimeDate AS DATE) as [Date], COUNT(Distinct SKU_ID) as [No of Orders], CL_ClientName as Client, ROUND((SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),1) as float))/3600)), 2) as [Total Time] " +
                        ", ROUND((COUNT(Distinct SKU_ID)/(SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),1) as float))/3600))) ,2) as [Effective Rate] " +
                        "from RTM_IPVDetails left join RTM_Records on RTM_IPVDetails.R_Id = RTM_Records.R_ID " +
                        "left join RTM_Client_List on RTM_Records.R_Client = RTM_Client_List.CL_ID " +
                        "Where team_id =1 and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, TimeDate))) ='" + indianTime.AddDays(-1).ToShortDateString() + "' group by CL_ClientName, CAST(TimeDate AS DATE) order by CL_ClientName";

            using (da = new SqlDataAdapter(sQuery, con))
            {
                da.Fill(dt);
            }

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow drRow in dt.Rows)
                {
                    dr = dtResult.NewRow();
                    dr["Date"] = drRow["Date"];
                    dr["Client"] = drRow["Client"];
                    dr["No of Orders"] = drRow["No of Orders"];
                    totalCount = totalCount + Convert.ToDouble(drRow["No of Orders"]);
                    dr["Total Time"] = drRow["Total Time"];
                    totaltime = totaltime + Convert.ToDouble(drRow["Total Time"]);
                    dr["Effective Rate"] = drRow["Effective Rate"];

                    dtResult.Rows.Add(dr);
                }

                dr = dtResult.NewRow();

                dr["Date"] = "";
                dr["Client"] = "Total";
                dr["No of Orders"] = Math.Round(totalCount, 2, MidpointRounding.AwayFromZero);
                dr["Total Time"] = Math.Round(totaltime, 2, MidpointRounding.AwayFromZero);
                dr["Effective Rate"] = Math.Round((totalCount / totaltime), 2, MidpointRounding.AwayFromZero);

                dtResult.Rows.Add(dr);

                getDelayHTML(dtResult);

                //string filePath = "" + Directory.GetCurrentDirectory() + "\\Effective Rate Fulfillment Projects - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv";
                //CSVUtility.ToCSV(dtResult, filePath);

                StringBuilder sb = new StringBuilder();

                sb.AppendLine("");
                sb.AppendLine(myBuilder.ToString());
                sb.AppendLine("");//here I want the data to       display in table format
                sb.AppendLine("");
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                sb.AppendLine("");

                MailMessage message1 = new MailMessage();

                DirectCSV csv = new DirectCSV();
                var data = csv.ExportToCSV(dtResult);
                var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(data);
                MemoryStream ms = new MemoryStream(bytes);

                Attachment attachFile = new Attachment(ms, "Effective Rate Fulfillment Projects - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv", "application/csv");
                message1.Attachments.Add(attachFile);

                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);

                message1.To.Add(new MailAddress("Sandeep.C@tangoe.com"));
                message1.CC.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));

                message1.Subject = "Effective Rate Fulfillment Projects - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy");

                message1.Body = sb.ToString();
                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment(filePath);
                //message1.Attachments.Add(attachment);
                message1.IsBodyHtml = true;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                //smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }
        }

        private DataTable getSKUDuration(string client, string date)
        {
            DataTable dtDuration = new DataTable();
            //string query = "Select COALESCE(A.IPtime,0) as IPtime, COALESCE(B.QCtime,0) as QCtime from " +
            //                "(select CL_ClientName, (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600)) as [IPtime] from RTM_Records WITH (NOLOCK) left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where CL_ClientName='"+ client +"' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) ='"+ date +"' and R_SubTask = '1740' Group by CL_ClientName) A " +
            //                "Left join "+
            //                "(select CL_ClientName, (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600)) as [QCtime] from RTM_Records WITH (NOLOCK) left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where CL_ClientName='" + client + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) ='" + date + "' and R_SubTask = '213' Group by CL_ClientName) B " +
            //                 "on A.CL_ClientName = B.CL_ClientName";

            string query = "select " +
                            "(select COALESCE((SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600)),0) from RTM_Records WITH (NOLOCK) left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where CL_ClientName='" + client + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) ='" + date + "' and R_SubTask = '1740' Group by CL_ClientName) as [IPtime], " +
                            "(select COALESCE((SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600)),0) from RTM_Records WITH (NOLOCK) left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where CL_ClientName='" + client + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) ='" + date + "' and R_SubTask = '213' Group by CL_ClientName) as [QCtime] ";

            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dtDuration);
            }

            return dtDuration;
        }

        private DataTable getCMPDuration(string client, string date)
        {
            DataTable dtDuration = new DataTable();
            //string query = "Select COALESCE(A.IPtime,0) as IPtime, COALESCE(B.QCtime,0) as QCtime from " +
            //                "(select CL_ClientName, (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600)) as [IPtime] from RTM_Records WITH (NOLOCK) left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where CL_ClientName='" + client + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) ='" + date + "' and (R_SubTask = '1694' or R_SubTask = '11105') Group by CL_ClientName) A " +
            //                "Left join " +
            //                "(select CL_ClientName, (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600)) as [QCtime] from RTM_Records WITH (NOLOCK) left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where CL_ClientName='" + client + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) ='" + date + "' and (R_SubTask = '202' or R_SubTask = '11106') Group by CL_ClientName) B " +
            //                 "on A.CL_ClientName = B.CL_ClientName";

            string query = "select " +
                            "(select COALESCE((SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600)),0) from RTM_Records WITH (NOLOCK) left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where CL_ClientName='" + client + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) ='" + date + "' and (R_SubTask = '1694' or R_SubTask = '11105') Group by CL_ClientName) as [IPtime], " +
                            "(select COALESCE((SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600)),0) from RTM_Records WITH (NOLOCK) left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where CL_ClientName='" + client + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) ='" + date + "' and (R_SubTask = '202' or R_SubTask = '11106') Group by CL_ClientName) as [QCtime] ";
            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dtDuration);
            }

            return dtDuration;
        }

        private DataTable getSKUWeekDuration(string client, string from, string to)
        {
            DataTable dtDuration = new DataTable();
            string query = "Select COALESCE(A.IPtime,0) as IPtime, COALESCE(B.QCtime,0) as QCtime from " +
                            "(select CL_ClientName, (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600)) as [IPtime] from RTM_Records WITH (NOLOCK) left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where CL_ClientName='" + client + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) Between '" + from + "' and '" + to + "' and R_SubTask = '1740' Group by CL_ClientName) A " +
                            "Left join " +
                            "(select CL_ClientName, (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600)) as [QCtime] from RTM_Records WITH (NOLOCK) left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where CL_ClientName='" + client + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) Between '" + from + "' and '" + to + "' and R_SubTask = '213' Group by CL_ClientName) B " +
                             "on A.CL_ClientName = B.CL_ClientName";
            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dtDuration);
            }

            return dtDuration;
        }

        //Effective Rate  Weekly Report- Invoice Team
        private void SendWeeklySKUEffectiveRateFromSKUDB()
        {

            DataRow dr;
            BuildSKUDBERTable();
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            dt = new DataTable();
            string sQuery = "";
            //get clients details
            sQuery = "select COUNT(SKU_NUMBER) as [SKU], CL_CLientName as Client from RTM_Sku " +
                        "left join RTM_Client_List on TSHEETS_CLIENT_CODE = CL_Code " +
                        "left join RTM_User_List on fullname = RIGHT(UL_System_User_Name, LEN(UL_System_User_Name) - 5) " +
                        "WHERE CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, DATE_FINISHED)))  Between '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 7).ToShortDateString() + "' and '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToShortDateString() + "' and CL_TeamId = 29 and UL_Team_ID = 29 and CL_Product = 'IPV' " +
                        "Group by CL_ClientName";

            using (da = new SqlDataAdapter(sQuery, con))
            {
                da.Fill(dt);
            }
            double totalIPDuration = 0;
            double totalQCDuration = 0;
            int skuCount = 0;
            //
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow drRow in dt.Rows)
                {
                    dr = dtResult.NewRow();
                    dr["Date"] = indianTime.AddDays(-(int)indianTime.DayOfWeek - 7).ToShortDateString() + " - " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToShortDateString();
                    dr["Client Name"] = drRow["Client"];
                    dr["No of Invoices"] = drRow["SKU"];
                    skuCount = skuCount + Convert.ToInt32(drRow["SKU"]);
                    DataTable dtDuration = new DataTable();
                    //Get weekly Report Data************
                    dtDuration = getSKUWeekDuration(drRow["Client"].ToString(), indianTime.AddDays(-(int)indianTime.DayOfWeek - 7).ToShortDateString(), indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToShortDateString());
                    double totalIPDur = 0;
                    double totalQCDur = 0;
                    if (dtDuration.Rows.Count > 0)
                    {
                        totalIPDur = Math.Round(Convert.ToDouble(dtDuration.Rows[0]["IPtime"]), 2, MidpointRounding.AwayFromZero);
                        totalIPDuration = totalIPDuration + totalIPDur;

                        totalQCDur = Math.Round(Convert.ToDouble(dtDuration.Rows[0]["QCtime"]), 2, MidpointRounding.AwayFromZero);
                        totalQCDuration = totalQCDuration + totalQCDur;
                    }
                    else
                    {
                        totalIPDur = 0;
                        totalQCDur = 0;
                    }

                    dr["IP Time"] = totalIPDur;
                    dr["QC Time"] = totalQCDur;
                    double totalTime = totalIPDur + totalQCDur;
                    dr["Total Time"] = Math.Round((totalIPDur + totalQCDur), 2, MidpointRounding.AwayFromZero);
                    if (totalIPDur > 0)
                    {
                        dr["Effective Rate - IP"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SKU"]) / totalIPDur), 2, MidpointRounding.AwayFromZero);
                    }
                    else
                    {
                        dr["Effective Rate - IP"] = "";
                    }

                    if (totalQCDur > 0)
                    {
                        dr["Effective Rate - QC"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SKU"]) / totalQCDur), 2, MidpointRounding.AwayFromZero);
                    }
                    else
                    {
                        dr["Effective Rate - QC"] = "";
                    }

                    if (totalTime > 0)
                    {
                        dr["Effective Rate - Overall"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SKU"]) / totalTime), 2, MidpointRounding.AwayFromZero);
                    }
                    else
                    {
                        dr["Effective Rate - Overall"] = "";
                    }

                    dtResult.Rows.Add(dr);

                }
                dr = dtResult.NewRow();
                dr["Date"] = "";
                dr["Client Name"] = "Total";
                dr["No of Invoices"] = skuCount;
                dr["IP Time"] = totalIPDuration;
                dr["QC Time"] = totalQCDuration;
                dr["Total Time"] = Math.Round((totalIPDuration + totalQCDuration), 2, MidpointRounding.AwayFromZero);
                dr["Effective Rate - IP"] = Math.Round(skuCount / totalIPDuration, 2, MidpointRounding.AwayFromZero);
                dr["Effective Rate - QC"] = Math.Round(skuCount / totalQCDuration, 2, MidpointRounding.AwayFromZero);
                double grandTotal = Math.Round((totalIPDuration + totalQCDuration), 2, MidpointRounding.AwayFromZero);
                dr["Effective Rate - Overall"] = Math.Round(Convert.ToDouble(skuCount / grandTotal), 2, MidpointRounding.AwayFromZero);
                dtResult.Rows.Add(dr);


            }


            if (dtResult.Rows.Count > 0)
            {
                getDelayHTML(dtResult);

                //string filePath = "" + Directory.GetCurrentDirectory() + "\\Weekly Effective Rate Invoices - From" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 7).ToString("MM-dd-yyyy") + " to " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToString("MM-dd-yyyy") + ".csv";
                //CSVUtility.ToCSV(dtResult, filePath);

                StringBuilder sb = new StringBuilder();

                sb.AppendLine("");
                sb.AppendLine(myBuilder.ToString());
                sb.AppendLine("");//here I want the data to       display in table format
                sb.AppendLine("");
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                sb.AppendLine("");

                MailMessage message1 = new MailMessage();

                DirectCSV csv = new DirectCSV();
                var data = csv.ExportToCSV(dtResult);
                var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(data);
                MemoryStream ms = new MemoryStream(bytes);

                Attachment attachFile = new Attachment(ms, "Weekly Effective Rate IPV Invoices - From" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 7).ToString("MM-dd-yyyy") + " to " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToString("MM-dd-yyyy") + ".csv", "application/csv");
                message1.Attachments.Add(attachFile);

                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);

                message1.To.Add(new MailAddress("Sriram.Krishnan@tangoe.com"));
                message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                message1.To.Add(new MailAddress("Sandesh.Ravichandra@tangoe.com"));                
                message1.CC.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));           
                message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));

              
                message1.Subject = "Weekly Effective Rate IPV Invoices - From" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 7).ToString("MM-dd-yyyy") + " to " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToString("MM-dd-yyyy");

                message1.Body = sb.ToString();
                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment(filePath);
                //message1.Attachments.Add(attachment);
                message1.IsBodyHtml = true;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                //smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }
        }

        //Effective Rate Daily Report - Invoices Team
        private void SendWeeklySKUEffectiveRateFromSKUDBDay()
        {

            DataRow dr;
            BuildSKUDBERTable();
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            dt = new DataTable();
            string sQuery = "";
            //get clients details
            sQuery = "select COUNT(SKU_NUMBER) as [SKU], CL_CLientName as Client from RTM_Sku " +
                        "left join RTM_Client_List on TSHEETS_CLIENT_CODE = CL_Code " +
                        "left join RTM_User_List on fullname = RIGHT(UL_System_User_Name, LEN(UL_System_User_Name) - 5) " +
                        "WHERE CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, DATE_FINISHED)))  Between '" + indianTime.AddDays(-1).ToShortDateString() + "' and '" + indianTime.AddDays(-1).ToShortDateString() + "' and CL_TeamId = 29 and UL_Team_ID = 29 and CL_Product = 'IPV' " +
                        "Group by CL_ClientName";

            using (da = new SqlDataAdapter(sQuery, con))
            {
                da.Fill(dt);
            }


            double totalIPDurDaily = 0;
            double totalQCDurDaily = 0;
            double totalIPDurationDaily = 0;
            double totalQCDurationDaily = 0;
            int skuCountDaily = 0;

            //
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow drRow in dt.Rows)
                {
                    //Get Daily Report *************      

                    dr = dtResult.NewRow();
                    string str = indianTime.AddDays(-1).ToShortDateString() + " - " + indianTime.AddDays(-1).ToShortDateString();
                    dr["Date"] = indianTime.AddDays(-1).ToShortDateString() + " - " + indianTime.AddDays(-1).ToShortDateString();
                    dr["Client Name"] = drRow["Client"];
                    dr["No of Invoices"] = drRow["SKU"];
                    skuCountDaily = skuCountDaily + Convert.ToInt32(drRow["SKU"]);

                    //Get Daily Report Data************
                    DataTable dtDuration_Daily = new DataTable();
                    //dtDuration_Daily = getSKUWeekDuration(drRow["Client"].ToString(), indianTime.AddDays(-(int)indianTime.DayOfWeek - 7).ToShortDateString(), indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToShortDateString());

                    dtDuration_Daily = getSKUWeekDuration(drRow["Client"].ToString(), indianTime.AddDays(-1).ToShortDateString(), indianTime.AddDays(-1).ToShortDateString());

                    if (dtDuration_Daily.Rows.Count > 0)
                    {
                        totalIPDurDaily = Math.Round(Convert.ToDouble(dtDuration_Daily.Rows[0]["IPtime"]), 2, MidpointRounding.AwayFromZero);
                        totalIPDurationDaily = totalIPDurationDaily + totalIPDurDaily;

                        totalQCDurDaily = Math.Round(Convert.ToDouble(dtDuration_Daily.Rows[0]["QCtime"]), 2, MidpointRounding.AwayFromZero);
                        totalQCDurationDaily = totalQCDurationDaily + totalQCDurDaily;
                    }
                    else
                    {
                        totalIPDurDaily = 0;
                        totalQCDurDaily = 0;
                    }

                    dr["IP Time"] = totalIPDurDaily;
                    dr["QC Time"] = totalQCDurDaily;
                    double totalTimeDaily = totalIPDurDaily + totalQCDurDaily;
                    dr["Total Time"] = Math.Round((totalIPDurDaily + totalQCDurDaily), 2, MidpointRounding.AwayFromZero);
                    if (totalIPDurDaily > 0)
                    {
                        dr["Effective Rate - IP"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SKU"]) / totalIPDurDaily), 2, MidpointRounding.AwayFromZero);
                    }
                    else
                    {
                        dr["Effective Rate - IP"] = "";
                    }

                    if (totalQCDurDaily > 0)
                    {
                        dr["Effective Rate - QC"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SKU"]) / totalQCDurDaily), 2, MidpointRounding.AwayFromZero);
                    }
                    else
                    {
                        dr["Effective Rate - QC"] = "";
                    }

                    if (totalTimeDaily > 0)
                    {
                        dr["Effective Rate - Overall"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SKU"]) / totalTimeDaily), 2, MidpointRounding.AwayFromZero);
                    }
                    else
                    {
                        dr["Effective Rate - Overall"] = "";
                    }

                    dtResult.Rows.Add(dr);


                }

                dr = dtResult.NewRow();
                dr["Date"] = "";
                dr["Client Name"] = "Total";
                dr["No of Invoices"] = skuCountDaily;
                dr["IP Time"] = totalIPDurationDaily;
                dr["QC Time"] = totalQCDurationDaily;
                dr["Total Time"] = Math.Round((totalIPDurationDaily + totalQCDurationDaily), 2, MidpointRounding.AwayFromZero);
                dr["Effective Rate - IP"] = Math.Round(skuCountDaily / totalIPDurationDaily, 2, MidpointRounding.AwayFromZero);
                dr["Effective Rate - QC"] = Math.Round(skuCountDaily / totalQCDurationDaily, 2, MidpointRounding.AwayFromZero);
                double grandTotal = Math.Round((totalIPDurationDaily + totalQCDurationDaily), 2, MidpointRounding.AwayFromZero);
                dr["Effective Rate - Overall"] = Math.Round(Convert.ToDouble(skuCountDaily / grandTotal), 2, MidpointRounding.AwayFromZero);
                dtResult.Rows.Add(dr);


            }


            if (dtResult.Rows.Count > 0)
            {
                getDelayHTML(dtResult);

                StringBuilder sb = new StringBuilder();

                sb.AppendLine("");
                sb.AppendLine(myBuilder.ToString());
                sb.AppendLine("");//here I want the data to       display in table format
                sb.AppendLine("");
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                sb.AppendLine("");

                MailMessage message1 = new MailMessage();

                DirectCSV csv = new DirectCSV();
                var data = csv.ExportToCSV(dtResult);
                var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(data);
                MemoryStream ms = new MemoryStream(bytes);

                Attachment attachFile = new Attachment(ms, "Daily Effective Rate IPV Invoices - From" + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + " to " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + ".csv", "application/csv");
                message1.Attachments.Add(attachFile);

                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);

                message1.To.Add(new MailAddress("Sriram.Krishnan@tangoe.com"));
                message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                message1.To.Add(new MailAddress("Sandesh.Ravichandra@tangoe.com"));
                message1.CC.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));
              //  message1.To.Add(new MailAddress("namohar.m@tangoe.com"));

                message1.Subject = "Daily Effective Rate IPV Invoices - From " + indianTime.AddDays(-1).ToString("MM-dd-yyyy") + " to " + indianTime.AddDays(-1).ToString("MM-dd-yyyy");

                message1.Body = sb.ToString();
                message1.IsBodyHtml = true;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                //smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }
        }

        //Effective Rate Monthly Report - Invoices Team
        private void SendWeeklySKUEffectiveRateFromSKUDBMonthly()
        {
            double totalIPDurationMonthly = 0;
            double totalQCDurationMonthly = 0;
            int skuCountMonthly = 0;
            double totalIPDurMonthly = 0;
            double totalQCDurMonthly = 0;

            //  DateTime now = DateTime.Now;
            DateTime now = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            DateTime From = new DateTime(now.Year, now.Month - 1, 1);
            var DaysInMonth = DateTime.DaysInMonth(now.Year, now.Month - 1);
            DateTime To = new DateTime(now.Year, now.Month - 1, DaysInMonth);


            DataRow dr;
            BuildSKUDBERTable();
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            dt = new DataTable();
            string sQuery = "";
            //get clients details
            sQuery = "select COUNT(SKU_NUMBER) as [SKU], CL_CLientName as Client from RTM_Sku " +
                        "left join RTM_Client_List on TSHEETS_CLIENT_CODE = CL_Code " +
                        "left join RTM_User_List on fullname = RIGHT(UL_System_User_Name, LEN(UL_System_User_Name) - 5) " +
                        "WHERE CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, DATE_FINISHED)))  Between '" + From.ToShortDateString() + "' and '" + To.ToShortDateString() + "' and CL_TeamId = 29 and UL_Team_ID = 29 and CL_Product = 'IPV' " +
                        "Group by CL_ClientName";

            using (da = new SqlDataAdapter(sQuery, con))
            {
                da.Fill(dt);
            }

            //
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow drRow in dt.Rows)
                {

                    dr = dtResult.NewRow();
                    dr["Date"] = From.ToShortDateString() + " - " + To.ToShortDateString();
                    dr["Client Name"] = drRow["Client"];
                    dr["No of Invoices"] = drRow["SKU"];
                    skuCountMonthly = skuCountMonthly + Convert.ToInt32(drRow["SKU"]);
                    DataTable dtDurationMonthly = new DataTable();

                    //Get Monthly Report Data************
                    dtDurationMonthly = getSKUWeekDuration(drRow["Client"].ToString(), From.ToShortDateString(), To.ToShortDateString());

                    if (dtDurationMonthly.Rows.Count > 0)
                    {
                        totalIPDurMonthly = Math.Round(Convert.ToDouble(dtDurationMonthly.Rows[0]["IPtime"]), 2, MidpointRounding.AwayFromZero);
                        totalIPDurationMonthly = totalIPDurationMonthly + totalIPDurMonthly;

                        totalQCDurMonthly = Math.Round(Convert.ToDouble(dtDurationMonthly.Rows[0]["QCtime"]), 2, MidpointRounding.AwayFromZero);
                        totalQCDurationMonthly = totalQCDurationMonthly + totalQCDurMonthly;
                    }
                    else
                    {
                        totalIPDurMonthly = 0;
                        totalQCDurMonthly = 0;
                    }

                    dr["IP Time"] = totalIPDurMonthly;
                    dr["QC Time"] = totalQCDurMonthly;
                    double totalTimeMonthly = totalIPDurMonthly + totalQCDurMonthly;
                    dr["Total Time"] = Math.Round((totalIPDurMonthly + totalQCDurMonthly), 2, MidpointRounding.AwayFromZero);
                    if (totalIPDurMonthly > 0)
                    {
                        dr["Effective Rate - IP"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SKU"]) / totalIPDurMonthly), 2, MidpointRounding.AwayFromZero);
                    }
                    else
                    {
                        dr["Effective Rate - IP"] = "";
                    }

                    if (totalQCDurMonthly > 0)
                    {
                        dr["Effective Rate - QC"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SKU"]) / totalQCDurMonthly), 2, MidpointRounding.AwayFromZero);
                    }
                    else
                    {
                        dr["Effective Rate - QC"] = "";
                    }

                    if (totalTimeMonthly > 0)
                    {
                        dr["Effective Rate - Overall"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SKU"]) / totalTimeMonthly), 2, MidpointRounding.AwayFromZero);
                    }
                    else
                    {
                        dr["Effective Rate - Overall"] = "";
                    }

                    dtResult.Rows.Add(dr);
                }

                dr = dtResult.NewRow();
                dr["Date"] = "";
                dr["Client Name"] = "Total";
                dr["No of Invoices"] = skuCountMonthly;
                dr["IP Time"] = totalIPDurationMonthly;
                dr["QC Time"] = totalQCDurationMonthly;
                dr["Total Time"] = Math.Round((totalIPDurationMonthly + totalQCDurationMonthly), 2, MidpointRounding.AwayFromZero);
                dr["Effective Rate - IP"] = Math.Round(skuCountMonthly / totalIPDurationMonthly, 2, MidpointRounding.AwayFromZero);
                dr["Effective Rate - QC"] = Math.Round(skuCountMonthly / totalQCDurationMonthly, 2, MidpointRounding.AwayFromZero);
                double grandTotal = Math.Round((totalIPDurationMonthly + totalQCDurationMonthly), 2, MidpointRounding.AwayFromZero);
                dr["Effective Rate - Overall"] = Math.Round(Convert.ToDouble(skuCountMonthly / grandTotal), 2, MidpointRounding.AwayFromZero);
                dtResult.Rows.Add(dr);


            }


            if (dtResult.Rows.Count > 0)
            {
                getDelayHTML(dtResult);

                //string filePath = "" + Directory.GetCurrentDirectory() + "\\Weekly Effective Rate Invoices - From" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 7).ToString("MM-dd-yyyy") + " to " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToString("MM-dd-yyyy") + ".csv";
                //CSVUtility.ToCSV(dtResult, filePath);

                StringBuilder sb = new StringBuilder();

                sb.AppendLine("");
                sb.AppendLine(myBuilder.ToString());
                sb.AppendLine("");//here I want the data to       display in table format
                sb.AppendLine("");
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                sb.AppendLine("");

                MailMessage message1 = new MailMessage();

                DirectCSV csv = new DirectCSV();
                var data = csv.ExportToCSV(dtResult);
                var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(data);
                MemoryStream ms = new MemoryStream(bytes);

                Attachment attachFile = new Attachment(ms, "Monthly Effective Rate IPV Invoices - From " + From.ToString("MM-dd-yyyy") + " to " + To.ToString("MM-dd-yyyy") + ".csv", "application/csv");
                message1.Attachments.Add(attachFile);

                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);

                message1.To.Add(new MailAddress("Sriram.Krishnan@tangoe.com"));
                message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                message1.To.Add(new MailAddress("Sandesh.Ravichandra@tangoe.com"));
                message1.CC.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));

               // message1.To.Add(new MailAddress("namohar.m@tangoe.com"));
                message1.Subject = "Monthly Effective Rate IPV Invoices - From " + From.ToString("MM-dd-yyyy") + " to " + To.ToString("MM-dd-yyyy");

                message1.Body = sb.ToString();
                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment(filePath);
                //message1.Attachments.Add(attachment);
                message1.IsBodyHtml = true;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                //smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;
                smtp.Send(message1);

            }
        }

        private void SendCMPDailyER_QC()
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            myBuilder = new StringBuilder();
            DataTable dtER = new DataTable();
            clsEffectiveRate objER = new clsEffectiveRate();
            dtER = objER.QCER_CMPDB_Client();
            if (dtER.Rows.Count > 0)
            {
                getHTMLTableQC(dtER);
            }
            dtER = new DataTable();
            dtER = objER.QCER_CMPDB_User();
            if (dtER.Rows.Count > 0)
            {
                getHTMLTableQC(dtER);
            }
            if (!string.IsNullOrEmpty(myBuilder.ToString()))
            {
                StringBuilder sb = new StringBuilder();

                sb.AppendLine("");
                sb.AppendLine(myBuilder.ToString());
                sb.AppendLine("");//here I want the data to       display in table format
                sb.AppendLine("");
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                sb.AppendLine("");

                MailMessage message1 = new MailMessage();
                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);

                //message1.To.Add(new MailAddress("Sriram.Krishnan@tangoe.com"));
                message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                // message1.To.Add(new MailAddress("Sandesh.Ravichandra@tangoe.com"));
                message1.To.Add(new MailAddress("Johwessly.Chennaiah@tangoe.com"));
                // message1.CC.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                // message1.CC.Add(new MailAddress("rich.lena@tangoe.com"));
                // message1.CC.Add(new MailAddress("melissa.guarracino@tangoe.com"));
                message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));

                message1.Subject = "CMP Effective Rate - Quality Check - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy");

                message1.Body = sb.ToString();
                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment(filePath);
                //message1.Attachments.Add(attachment);
                message1.IsBodyHtml = true;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                //smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }
        }

        private void SendSKUDailyER_QC()
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            myBuilder = new StringBuilder();
            DataTable dtER = new DataTable();
            clsEffectiveRate objER = new clsEffectiveRate();

            dtER = objER.QCER_SKUDB_CLIENT();
            if (dtER.Rows.Count > 0)
            {
                getHTMLTableQC(dtER);
            }
            dtER = new DataTable();

            dtER = objER.QCER_SKUDB_User();

            if (dtER.Rows.Count > 0)
            {
                getHTMLTableQC(dtER);
            }

            if (!string.IsNullOrEmpty(myBuilder.ToString()))
            {
                StringBuilder sb = new StringBuilder();

                sb.AppendLine("");
                sb.AppendLine(myBuilder.ToString());
                sb.AppendLine("");//here I want the data to       display in table format
                sb.AppendLine("");
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                sb.AppendLine("");

                MailMessage message1 = new MailMessage();
                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);


                message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                message1.To.Add(new MailAddress("Johwessly.Chennaiah@tangoe.com"));
                message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));

                //message1.To.Add(new MailAddress("namohar.m@tangoe.com"));
                //message1.CC.Add(new MailAddress("Namohar.M@tangoe.com"));
                message1.Subject = "IPV Effective Rate - Quality Check - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy");

                message1.Body = sb.ToString();
                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment(filePath);
                //message1.Attachments.Add(attachment);
                message1.IsBodyHtml = true;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                //smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }
        }

        public string getHTMLTableQC(DataTable dt)
        {
            //myBuilder = new StringBuilder();
            myBuilder.Append("<br />");
            myBuilder.Append("<table border='1' cellpadding='5' cellspacing='0'");
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
                if (myRow.ItemArray.Contains("Total"))
                {
                    myBuilder.Append("<tr align='left' valign='top'>");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        myBuilder.Append("<td align='left' valign='top' bgcolor='#FFFF00'>");
                        myBuilder.Append("<B />" + myRow[myColumn.ColumnName].ToString());
                        myBuilder.Append("</td>");
                    }
                    myBuilder.Append("</tr>");
                }
                else
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

            }
            myBuilder.Append("</table>");

            return myBuilder.ToString();
        }

        private void SendSubmitStatusToRashmi()
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            dt = new DataTable();
            string sQuery = "";

            sQuery = "select ROW_NUMBER() over (order by UL_User_Name) as [Sl.No.], UL_User_Name as [Employee Name], CASE R_Submit WHEN 1 THEN CONVERT(VARCHAR, R_SubmittedOn, 101)  ELSE 'Not submitted' END as [Submitted Date] from RTM_User_List WITH (NOLOCK) left join RTM_Records WITH (NOLOCK) on UL_User_Name = R_User_Name " +
                     "where (UL_RepMgrId = '102651' or UL_User_Name ='Deepak M') and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) BETWEEN '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 7).ToShortDateString() + "' and '" + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToShortDateString() + "' " +
                     "group by UL_User_Name, R_Submit, R_SubmittedOn order by UL_User_Name";

            using (da = new SqlDataAdapter(sQuery, con))
            {
                da.Fill(dt);
            }

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

                message1.From = new MailAddress(FromAddress);

                message1.CC.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));

                message1.Subject = "TimeSheet Submission status - " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 7).ToShortDateString() + " - " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToShortDateString();

                message1.Body = sb.ToString();

                message1.IsBodyHtml = true;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                //smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }
        }

        private void WeeklyTrackingReport()
        {
            clsWeeklyHours objHours = new clsWeeklyHours();
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            dt = new DataTable();

            string from = indianTime.AddDays(-(int)indianTime.DayOfWeek - 7).ToShortDateString();
            string to = indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToShortDateString();

            MailMessage message1 = new MailMessage();
            SmtpClient smtp = new SmtpClient();

            message1.From = new MailAddress(FromAddress);

            //message1.From = new MailAddress("Lokesha.B@tangoe.com");
            message1.To.Add(new MailAddress("AhujaDirects@tangoe.com"));
            message1.CC.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
            message1.CC.Add(new MailAddress("Shabeenaz1@tangoe.com"));
            message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));
            message1.Subject = "RTM Weekly Tracking Report for Less than 38.75 and more than 45 hours";

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Hi all,");
            sb.AppendLine("Please find report attached having the details of employees for whom time captured in less than 38.75 hrs. and more than 45 hours for the timesheet period<< " + from + " - " + to + " >> ");
            sb.AppendLine("");
            sb.AppendLine("");
            sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");

            message1.Body = sb.ToString();

            DirectExcel excel = new DirectExcel();
            DirectCSV csv = new DirectCSV();
            dt = objHours.GetWeeklyHoursLessThan38(from, to, "IND");
            var data = csv.ExportToCSV(dt);
            var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(data);
            MemoryStream ms = new MemoryStream(bytes);
            //System.IO.MemoryStream ms = excel.ExportToStream(dt);
            Attachment attachFile = new Attachment(ms, "Weekly_RTM_Hours_Less_Than_38.75.csv", "application/csv");
            message1.Attachments.Add(attachFile);

            dt = new DataTable();
            dt = objHours.GetWeeklyHoursGreaterThan45(from, to, "IND");

            var result = csv.ExportToCSV(dt);
            var bytes2 = Encoding.GetEncoding("iso-8859-1").GetBytes(result);

            MemoryStream stream = new MemoryStream(bytes2);

            attachFile = new Attachment(stream, "Weekly_RTM_Hours_Greater_Than_45.csv", "application/csv");

            message1.Attachments.Add(attachFile);

            //SmtpClient smtpClient = new SmtpClient("mail.north.tangoe.com");
            //smtpClient.UseDefaultCredentials = false;
            //NetworkCredential credentials = new NetworkCredential("Lokesha.B", "Lokeshmca11");
            //smtpClient.Credentials = credentials;
            //smtpClient.Send(message1);

            smtp.Port = 25;
            smtp.Host = "10.0.5.104";
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.EnableSsl = false;

            smtp.Send(message1);
        }

        private void ER_IN_Provisioning()
        {
            clsEffectiveRate objER = new clsEffectiveRate();
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            dt = new DataTable();
            dt = objER.GetERFromRTM(indianTime.AddDays(-1).ToShortDateString(), 21);

            if (dt.Rows.Count > 0)
            {
                MailMessage message1 = new MailMessage();
                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);
                message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));

                getDelayHTML(dt);

                StringBuilder sb = new StringBuilder();

                sb.AppendLine("");
                sb.AppendLine(myBuilder.ToString());
                sb.AppendLine("");//here I want the data to       display in table format
                sb.AppendLine("");
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                sb.AppendLine("");

                message1.Subject = "Effective Rate IN_Provisioning - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy");

                message1.Body = sb.ToString();

                message1.IsBodyHtml = true;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }
        }

        private void ER_IN_Inventory_Management()
        {
            clsEffectiveRate objER = new clsEffectiveRate();
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            dt = new DataTable();
            dt = objER.GetERFromRTM(indianTime.AddDays(-1).ToShortDateString(), 24);

            if (dt.Rows.Count > 0)
            {
                MailMessage message1 = new MailMessage();
                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);
                message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));

                getDelayHTML(dt);

                StringBuilder sb = new StringBuilder();

                sb.AppendLine("");
                sb.AppendLine(myBuilder.ToString());
                sb.AppendLine("");//here I want the data to       display in table format
                sb.AppendLine("");
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                sb.AppendLine("");

                message1.Subject = "Effective Rate IN_Inventory_Management - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy");

                message1.Body = sb.ToString();

                message1.IsBodyHtml = true;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }
        }

        private void Invoices_OIR_Processing()
        {
            clsInvoice_OIR objOIR = new clsInvoice_OIR();
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            dt = new DataTable();

            dt = objOIR.GetOIRTaskDetails(29, indianTime.ToShortDateString());

            if (dt.Rows.Count > 0)
            {
                MailMessage message1 = new MailMessage();
                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);

                //message1.To.Add(new MailAddress("Sriram.Krishnan@tangoe.com"));
                message1.To.Add(new MailAddress("blr_asentinel@tangoe.onmicrosoft.com"));
                message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));

                message1.Subject = "Invoice_OIR Task and Invoice Processing Subtask Daily Processing Report - " + indianTime.ToShortDateString();

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Hi all,");
                sb.AppendLine("Please find report attached for IN_Invoices team having the details of the time spend for the task OIR and Subtask Invoice processing.");
                sb.AppendLine("");
                sb.AppendLine("");
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");

                message1.Body = sb.ToString();

                DirectExcel excel = new DirectExcel();
                DirectCSV csv = new DirectCSV();
                var data = csv.ExportToCSV(dt);
                var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(data);
                MemoryStream ms = new MemoryStream(bytes);
                Attachment attachFile = new Attachment(ms, "Invoice_OIR_InvoiceProcessing_Report-" + indianTime.ToString("MM-dd-yyyy") + ".csv", "application/csv");
                message1.Attachments.Add(attachFile);

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }
        }

        private void GenerateRPAReport()
        {
            dtResult = new DataTable();
            RPAReport objRPA = new RPAReport();
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            string from = indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToShortDateString();
            string to = indianTime.ToShortDateString();

            dtResult = objRPA.GetRPSTaskDetails(from, to);

            if (dtResult.Rows.Count > 0)
            {
                MailMessage message1 = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                message1.From = new MailAddress(FromAddress);
                message1.To.Add(new MailAddress("trodriguez@tangoe.com"));
                message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));

                message1.Subject = "Time spent on RPA task from " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 1).ToString("MM-dd-yyyy") + "To" + indianTime.ToString("MM-dd-yyyy");

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Hi,");
                sb.AppendLine("Please find report attached for time spent on RPA task.");
                sb.AppendLine("");
                sb.AppendLine("");
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");

                message1.Body = sb.ToString();

                DirectCSV csv = new DirectCSV();
                var data = csv.ExportToCSV(dtResult);
                var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(data);
                MemoryStream ms = new MemoryStream(bytes);
                Attachment attachFile = new Attachment(ms, "RPA - " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToString("MM-dd-yyyy") + "To" + indianTime.AddDays(-(int)indianTime.DayOfWeek).ToString("MM-dd-yyyy") + ".csv", "application/csv");
                message1.Attachments.Add(attachFile);
                //string filePath = @"\\Apollo\common\General Electric Corporate\Implementation\International Project Scope\RPA\Time Worked\RPA - " + indianTime.AddDays(-(int)indianTime.DayOfWeek - 6).ToString("MM-dd-yyyy") + "To" + indianTime.AddDays(-(int)indianTime.DayOfWeek).ToString("MM-dd-yyyy") + ".csv";
                //CSVUtility.ToCSV(dtResult, filePath);

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }
        }

        //*********************Namohar Code Started 23 july 2018 ****************
        private void SendPSLDailyER_QC()
        {
            DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
            myBuilder = new StringBuilder();
            DataTable dtER = new DataTable();
            clsEffectiveRatePSLDB objER = new clsEffectiveRatePSLDB();
            dtER = objER.QCER_PSLDB_Client();
            if (dtER.Rows.Count > 0)
            {
                getHTMLTableQC(dtER);
            }
            dtER = new DataTable();
            dtER = objER.QCER_PSLDB_User();
            if (dtER.Rows.Count > 0)
            {
                getHTMLTableQC(dtER);
            }
            if (!string.IsNullOrEmpty(myBuilder.ToString()))
            {
                StringBuilder sb = new StringBuilder();

                sb.AppendLine("");
                sb.AppendLine(myBuilder.ToString());
                sb.AppendLine("");//here I want the data to       display in table format
                sb.AppendLine("");
                sb.AppendLine("This is a system generated mail. Please send mail to RTM-Support@tangoe.com if you have any issues.");
                sb.AppendLine("");

                MailMessage message1 = new MailMessage();
                SmtpClient smtp = new SmtpClient();

                message1.From = new MailAddress(FromAddress);

                ////message1.To.Add(new MailAddress("Sriram.Krishnan@tangoe.com"));
                //message1.To.Add(new MailAddress("Sudeep.Siddaiah@tangoe.com"));
                 message1.To.Add(new MailAddress("Sandesh.Ravichandra@tangoe.com"));
                message1.To.Add(new MailAddress("Johwessly.Chennaiah@tangoe.com"));
                //// message1.CC.Add(new MailAddress("Rashmi.Ahuja@tangoe.com"));
                //// message1.CC.Add(new MailAddress("rich.lena@tangoe.com"));
                //// message1.CC.Add(new MailAddress("melissa.guarracino@tangoe.com"));
                message1.CC.Add(new MailAddress("RTM-Support@tangoe.com"));
               // message1.To.Add(new MailAddress("namohar.m@tangoe.com"));
                message1.Subject = "PSLDB Effective Rate - Quality Check - " + indianTime.AddDays(-1).ToString("MM-dd-yyyy");

                message1.Body = sb.ToString();
                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment(filePath);
                //message1.Attachments.Add(attachment);
                message1.IsBodyHtml = true;

                smtp.Port = 25;
                smtp.Host = "10.0.5.104";
                //smtp.Host = "outlook-south.tangoe.com";
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = false;

                smtp.Send(message1);
            }
        }

     

    }
}