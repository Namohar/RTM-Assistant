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
    public class clsEffectiveRate
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

        private void BuildCMPDBQCClientERTable()
        {
            dtResult = new DataTable();
            dtResult.Columns.Add("Date");
            dtResult.Columns.Add("Client Name");
            dtResult.Columns.Add("No of Invoices");
            dtResult.Columns.Add("QC Time");
            dtResult.Columns.Add("Effective Rate");
        }

        public DataTable QCER_CMPDB_Client()
        {
            try
            {
                DataRow dr;
                BuildCMPDBQCClientERTable();
                DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
                dt = new DataTable();
                string sQuery = "";

                sQuery = "select COUNT(invoiceSubId) as [SID], custAbbr as [Client Code], CL_ClientName as [Client] from RTM_CMP_QC WITH (NOLOCK) " +
                        "left join RTM_Client_List WITH (NOLOCK) on custAbbr = CL_Code " +
                        "where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, validatedDateBackend))) = '" + indianTime.AddDays(-1).ToShortDateString() + "' and CL_TeamId = 9 and CL_Status = 1 and CL_Product ='CMP' " +
                        "group by custAbbr, CL_ClientName";

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
                        dtDuration = getCMPDuration(drRow["Client"].ToString(), indianTime.AddDays(-1).ToShortDateString());
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

        private DataTable getCMPDuration(string client, string date)
        {
            DataTable dtDuration = new DataTable();
            string query = "select CL_ClientName, COALESCE((SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600)),0) as [QCtime] from RTM_Records WITH (NOLOCK) left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where CL_ClientName='" + client + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) ='" + date + "' and R_SubTask = '202' Group by CL_ClientName";
            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dtDuration);
            }

            return dtDuration;
        }

        private void BuildCMPDBQCUserERTable()
        {
            dtResult = new DataTable();
            dtResult.Columns.Add("Date");
            dtResult.Columns.Add("User");
            dtResult.Columns.Add("No of Invoices");
            dtResult.Columns.Add("QC Time");
            dtResult.Columns.Add("Effective Rate");
        }

        public DataTable QCER_CMPDB_User()
        {
            try
            {
                DataRow dr;
                BuildCMPDBQCUserERTable();
                DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
                dt = new DataTable();
                string sQuery = "";

                sQuery = "select COUNT(invoiceSubId) as [SID], UL_User_Name from RTM_CMP_QC WITH (NOLOCK) " +
                        "left join RTM_User_List WITH (NOLOCK) on validatedBy = RIGHT(UL_System_User_Name, LEN(UL_System_User_Name) - 5) " +
                        "where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, validatedDateBackend))) = '" + indianTime.AddDays(-1).ToShortDateString() + "' " +
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
                        dtDuration = getCMPDuration_User(drRow["UL_User_Name"].ToString(), indianTime.AddDays(-1).ToShortDateString());
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

        private DataTable getCMPDuration_User(string user, string date)
        {
            DataTable dtDuration = new DataTable();
            string query = "select UL_User_Name, COALESCE((SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600)),0) as [QCtime] from RTM_Records WITH (NOLOCK) left join RTM_User_List WITH (NOLOCK) on R_User_Name = UL_User_Name where UL_User_Name='" + user + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) ='" + date + "' and R_SubTask = '202' Group by UL_User_Name";
            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dtDuration);
            }

            return dtDuration;
        }

        private void BuildSKUDBQCClientERTable()
        {
            dtResult = new DataTable();
            dtResult.Columns.Add("Date");
            dtResult.Columns.Add("Client Name");
            dtResult.Columns.Add("No of QC Invoices");
            dtResult.Columns.Add("No of IP Invoices");
            dtResult.Columns.Add("QC Time");
            dtResult.Columns.Add("Effective Rate");
        }

        public DataTable QCER_SKUDB_CLIENT()
        {
            try
            {
                DataRow dr;
                BuildSKUDBQCClientERTable();
                DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
                dt = new DataTable();
                string sQuery = "";

                sQuery = "select COUNT(SKU_NUMBER) as [SKU], TSHEETS_CLIENT_CODE as [Client Code], CL_ClientName as [Client] from RTM_SKU_QC WITH (NOLOCK) " +
                        "left join RTM_Client_List WITH (NOLOCK) on TSHEETS_CLIENT_CODE = CL_Code " +
                        "where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, VALIDATED_DATETIME))) = '" + indianTime.AddDays(-1).ToShortDateString() + "' and CL_TeamId = 9 and CL_Status = 1 and CL_Product ='IPV' " +
                        "group by TSHEETS_CLIENT_CODE, CL_ClientName";

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
                        string clientCode = drRow["Client Code"].ToString();
                        dr["Date"] = indianTime.AddDays(-1).ToShortDateString();
                        dr["Client Name"] = drRow["Client"];
                        dr["No of QC Invoices"] = drRow["SKU"];
                        skuCount = skuCount + Convert.ToInt32(drRow["SKU"]);
                        DataTable dtIp = new DataTable();
                        dtIp = getIPSKUCount(clientCode, indianTime.AddDays(-1).ToShortDateString());
                        if (dtIp.Rows.Count > 0)
                        {
                            dr["No of IP Invoices"] = dtIp.Rows[0]["IPSKU"].ToString();
                        }
                        DataTable dtDuration = new DataTable();
                        dtDuration = getSKUDuration_Client(drRow["Client"].ToString(), indianTime.AddDays(-1).ToShortDateString());
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
                            dr["Effective Rate"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SKU"]) / totalQCDur), 2, MidpointRounding.AwayFromZero);
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
                    dr["No of QC Invoices"] = skuCount;
                    dr["No of IP Invoices"] = "";
                    dr["QC Time"] = totalQCDuration;
                    dr["Effective Rate"] = Math.Round(skuCount / totalQCDuration, 2, MidpointRounding.AwayFromZero);
                    dtResult.Rows.Add(dr);
                }
            }
            catch (Exception)
            {                
                
            }            

            return dtResult;
        }

        private DataTable getIPSKUCount(string clientCode, string date)
        {
            DataTable dtIP = new DataTable();
            string query = "select COUNT(SKU_NUMBER) as [IPSKU] from dbo.RTM_Sku where TSHEETS_CLIENT_CODE ='" + clientCode + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, DATE_FINISHED))) = '" + date + "'";
            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dtIP);
            }

            return dtIP;
        }

        private DataTable getSKUDuration_Client(string client, string date)
        {
            DataTable dtDuration = new DataTable();
            string query = "select CL_ClientName, (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600)) as [QCtime] from RTM_Records WITH (NOLOCK) left join RTM_Client_List WITH (NOLOCK) on R_Client = CL_ID where CL_ClientName='" + client + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) ='" + date + "' and R_SubTask = '213' Group by CL_ClientName";
            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dtDuration);
            }

            return dtDuration;
        }

        private void BuildSKUDBQCUserERTable()
        {
            dtResult = new DataTable();
            dtResult.Columns.Add("Date");
            dtResult.Columns.Add("User");
            dtResult.Columns.Add("No of Invoices");
            dtResult.Columns.Add("QC Time");
            dtResult.Columns.Add("Effective Rate");
        }

        public DataTable QCER_SKUDB_User()
        {
            try
            {
                DataRow dr;
                BuildSKUDBQCUserERTable();
                DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);
                dt = new DataTable();
                string sQuery = "";

                sQuery = "select COUNT(SKU_NUMBER) as [SQU], UL_User_Name from RTM_SKU_QC WITH (NOLOCK) "+
                        "left join RTM_User_List WITH (NOLOCK) on fullname = RIGHT(UL_System_User_Name, LEN(UL_System_User_Name) - 5) "+
                        "where CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, VALIDATED_DATETIME))) = '"+ indianTime.AddDays(-1).ToShortDateString() +"' "+
                        "group by UL_User_Name";                

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
                        dr["No of Invoices"] = drRow["SQU"];
                        skuCount = skuCount + Convert.ToInt32(drRow["SQU"]);
                        DataTable dtDuration = new DataTable();
                        dtDuration = getSKUDuration_User(drRow["UL_User_Name"].ToString(), indianTime.AddDays(-1).ToShortDateString());
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
                            dr["Effective Rate"] = Math.Round(Convert.ToDouble(Convert.ToInt32(drRow["SQU"]) / totalQCDur), 2, MidpointRounding.AwayFromZero);
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

        private DataTable getSKUDuration_User(string user, string date)
        {
            DataTable dtDuration = new DataTable();
            string query = "select UL_User_Name, COALESCE((SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(R_Duration,':','.'),1) as float))/3600)),0) as [QCtime] from RTM_Records WITH (NOLOCK) left join RTM_User_List WITH (NOLOCK) on R_User_Name = UL_User_Name where UL_User_Name='" + user + "' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) ='" + date + "' and R_SubTask = '213' Group by UL_User_Name";
            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dtDuration);
            }

            return dtDuration;
        }

        private void BuildRTMERTable()
        {
            dtResult = new DataTable();
            dtResult.Columns.Add("Date");
            dtResult.Columns.Add("Client Name");
            dtResult.Columns.Add("No of Invoices");
            dtResult.Columns.Add("Total Time");
            dtResult.Columns.Add("Effective Rate");
        }

        public DataTable GetERFromRTM(string date, int teamId)
        {
            dt = new DataTable();
            BuildRTMERTable();
            DataRow dr;
            string query = "select CL_CLientName as [Client Name], COUNT(Distinct SKU_Id) as [Count of Accounts], "+
                            "COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),1) as float))/3600),2),0) as [Total Hours], "+
                            "CASE WHEN COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),1) as float))/3600),2),0) = 0 THEN 0 ELSE ROUND((COUNT(Distinct SKU_Id) / COALESCE(ROUND(SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),3) as float)) + (SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),2) as float))/60) + (SUM(CAST(PARSENAME(REPLACE(Duration,':','.'),1) as float))/3600),2),0)),2) END as [Effective Rate] "+
                            " from dbo.RTM_IPVDetails as IP "+
                            "LEFT JOIN RTM_Records as R ON IP.R_Id = R.R_ID "+
                            "LEFT JOIN RTM_Client_List ON R_Client = CL_ID  "+
                            "where Team_Id = '"+ teamId +"' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, TimeDate))) = '"+ date +"' "+
                            "Group BY CL_CLientName order by CL_CLientName";

            using (da = new SqlDataAdapter(query, con))
            {
                da.Fill(dt);
            }

            double totalDuration = 0;
            int totalCount = 0;

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow drRow in dt.Rows)
                {
                    dr = dtResult.NewRow();

                    dr["Date"] = date;
                    dr["Client Name"] = drRow["Client Name"];
                    dr["No of Invoices"] = drRow["Count of Accounts"];
                    dr["Total Time"] = drRow["Total Hours"];
                    totalDuration = totalDuration + Math.Round(Convert.ToDouble(drRow["Total Hours"]), 2, MidpointRounding.AwayFromZero);
                    dr["Effective Rate"] = drRow["Effective Rate"];
                    totalCount = totalCount + Convert.ToInt32(drRow["Count of Accounts"]);
                    dtResult.Rows.Add(dr);
                }

                dr = dtResult.NewRow();
                dr["Date"] = "";
                dr["Client Name"] = "Total";
                dr["No of Invoices"] = totalCount;
                dr["Total Time"] = Math.Round(totalDuration, 2, MidpointRounding.AwayFromZero);
                dr["Effective Rate"] = Math.Round((totalCount / totalDuration), 2, MidpointRounding.AwayFromZero);

                dtResult.Rows.Add(dr);
            }

            return dtResult;
        }
    }
}
