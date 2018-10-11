using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;

namespace RTMReportsAssistant
{
    class clsEffectiveRatePSLDB
    {
        SqlConnection con = new SqlConnection(@"Data Source=10.55.5.40,1433;Initial Catalog=RTM_Global_Test;User ID=PRODRTMDB;Password=Prodrtm@123;");
        DataTable dtResult = new DataTable();
        DataTable dt = new DataTable();
        StringBuilder myBuilder = new StringBuilder();
        SqlDataAdapter da;

        string FromAddress = System.Configuration.ConfigurationManager.AppSettings["FromAddress"].ToString();
        // int port = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["FromAddress"]);
        string host = System.Configuration.ConfigurationManager.AppSettings["SMTPClient"].ToString();
        private static TimeZoneInfo INDIAN_ZONE = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");

        private void BuildPSLDBClientERTable()
        {
            dtResult = new DataTable();
            dtResult.Columns.Add("Date");
            dtResult.Columns.Add("Client Name");
            dtResult.Columns.Add("No of Invoices");
            dtResult.Columns.Add("QC Time");
            dtResult.Columns.Add("Effective Rate");
        }

        private void BuildPSLDBQCUserERTable()
        {
            dtResult = new DataTable();
            dtResult.Columns.Add("Date");
            dtResult.Columns.Add("User");
            dtResult.Columns.Add("No of Invoices");
            dtResult.Columns.Add("QC Time");
            dtResult.Columns.Add("Effective Rate");
        }

        public DataTable QCER_PSLDB_Client()
        {
            try
            {
                DataRow dr;             
                BuildPSLDBClientERTable();
                DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
                dt = new DataTable();
                string sQuery = "";
                sQuery = "select COUNT(distinct IPN) as [SID], TSHEETS_CLIENT_CODE as [Client Code], CL_ClientName as [Client] from RTM_PSL_QC WITH (NOLOCK) " +
                        "left join RTM_Client_List WITH (NOLOCK) on TSHEETS_CLIENT_CODE = CL_Code " +
                        "where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, ValidatedDateTime))) = '" + indianTime.AddDays(-1).ToShortDateString() + "' and CL_TeamId = 9 and CL_Status = 1 " +
                        "group by TSHEETS_CLIENT_CODE, CL_ClientName";

                //sQuery = "select COUNT(invoiceSubId) as [SID], custAbbr as [Client Code], CL_ClientName as [Client] from RTM_CMP " +
                //        "left join RTM_Client_List on custAbbr = CL_Code " +
                //        "where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, buildDateBackend))) = '09/20/2017' and CL_TeamId = 29 and CL_Status = 1 and CL_Product ='CMP' " +
                //        "group by custAbbr, CL_ClientName";

                using (da = new SqlDataAdapter(sQuery, con))
                {
                    da.Fill(dt);
                }
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
                        dtDuration = getPSLDurationClient(drRow["Client"].ToString(), indianTime.AddDays(-1).ToShortDateString());
                        //dtDuration = getCMPDuration(drRow["Client"].ToString(), "09/20/2017");                        
                        double totalQCDur = 0;
                        if (dtDuration.Rows.Count > 0)
                        {
                            totalQCDur = Math.Round(Convert.ToDouble(dtDuration.Rows[0]["QCtime"]), 2, MidpointRounding.AwayFromZero);
                            totalQCDuration = totalQCDuration + totalQCDur;
                        }
                        else
                        {
                            totalQCDur = 0;
                        }

                        dr["QC Time"] = totalQCDur;


                        if (totalQCDur > 0)
                        {
                            dr["Effective Rate"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SID"]) / totalQCDur), 2, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            dr["Effective Rate"] = "";
                        }

                        dtResult.Rows.Add(dr);
                    }

                    dr = dtResult.NewRow();

                    dr["Date"] = "";
                    dr["Client Name"] = "Total";
                    dr["No of Invoices"] = skuCount;
                    dr["QC Time"] = totalQCDuration;
                    dr["Effective Rate"] = Math.Round(skuCount / totalQCDuration, 2, MidpointRounding.AwayFromZero);
                    dtResult.Rows.Add(dr);
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }

            return dtResult;
        }

        private DataTable getPSLDurationClient(string client, string date)
        {
            DataTable dtDuration = new DataTable();
            string query = "select CL_ClientName, COALESCE((SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600)),0) as [QCtime] from RTM_Records WITH (NOLOCK) left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where CL_ClientName='" + client + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) ='" + date + "' and R_SubTask = '216' Group by CL_ClientName";
            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dtDuration);
            }

            return dtDuration;
        }


        public DataTable QCER_PSLDB_User()
        {
            try
            {
                DataRow dr;
                BuildPSLDBQCUserERTable();
                DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
                dt = new DataTable();
                string sQuery = "";

                sQuery = "select COUNT(distinct IPN) as [SID], UL_User_Name from RTM_PSL_QC WITH (NOLOCK) " +
                        "left join RTM_User_List WITH (NOLOCK) on fullname = RIGHT(UL_System_User_Name, LEN(UL_System_User_Name) - 5) " +
                        "where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, ValidatedDateTime))) = '" + indianTime.AddDays(-1).ToShortDateString() + "' " +
                        "group by UL_User_Name";

                //sQuery = "select COUNT(invoiceSubId) as [SID], custAbbr as [Client Code], CL_ClientName as [Client] from RTM_CMP " +
                //        "left join RTM_Client_List on custAbbr = CL_Code " +
                //        "where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, buildDateBackend))) = '09/20/2017' and CL_TeamId = 29 and CL_Status = 1 and CL_Product ='CMP' " +
                //        "group by custAbbr, CL_ClientName";

                using (da = new SqlDataAdapter(sQuery, con))
                {
                    da.Fill(dt);
                }
                double totalQCDuration = 0;
                int skuCount = 0;

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow drRow in dt.Rows)
                    {
                        dr = dtResult.NewRow();
                        dr["Date"] = indianTime.AddDays(-1).ToShortDateString();
                        dr["User"] = drRow["UL_User_Name"];
                        dr["No of Invoices"] = drRow["SID"];
                        skuCount = skuCount + Convert.ToInt32(drRow["SID"]);
                        DataTable dtDuration = new DataTable();
                        dtDuration = getPSLDuration_User(drRow["UL_User_Name"].ToString(), indianTime.AddDays(-1).ToShortDateString());
                        //dtDuration = getCMPDuration(drRow["Client"].ToString(), "09/20/2017");                        
                        double totalQCDur = 0;
                        if (dtDuration.Rows.Count > 0)
                        {
                            totalQCDur = Math.Round(Convert.ToDouble(dtDuration.Rows[0]["QCtime"]), 2, MidpointRounding.AwayFromZero);
                            totalQCDuration = totalQCDuration + totalQCDur;
                        }
                        else
                        {
                            totalQCDur = 0;
                        }

                        dr["QC Time"] = totalQCDur;


                        if (totalQCDur > 0)
                        {
                            dr["Effective Rate"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SID"]) / totalQCDur), 2, MidpointRounding.AwayFromZero);
                        }
                        else
                        {
                            dr["Effective Rate"] = "";
                        }

                        dtResult.Rows.Add(dr);
                    }

                    dr = dtResult.NewRow();

                    dr["Date"] = "";
                    dr["User"] = "Total";
                    dr["No of Invoices"] = skuCount;
                    dr["QC Time"] = totalQCDuration;
                    dr["Effective Rate"] = Math.Round(skuCount / totalQCDuration, 2, MidpointRounding.AwayFromZero);
                    dtResult.Rows.Add(dr);
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }

            return dtResult;
        }

        private DataTable getPSLDuration_User(string user, string date)
        {
            DataTable dtDuration = new DataTable();
            string query = "select UL_User_Name, COALESCE((SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600)),0) as [QCtime] from RTM_Records WITH (NOLOCK) left join RTM_User_List WITH (NOLOCK) on R_User_Name = UL_User_Name where UL_User_Name='" + user + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) ='" + date + "' and R_SubTask = '216' Group by UL_User_Name";
            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dtDuration);
            }

            return dtDuration;
        }
    }
}
