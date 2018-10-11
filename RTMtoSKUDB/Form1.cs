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
using System.Data.SqlTypes;

namespace RTMtoSKUDB
{
    public partial class Form1 : Form
    {
        SqlConnection con = new SqlConnection(@"Data Source=10.55.5.40,1433;Initial Catalog=RTM_Global_Test;User ID=PRODRTMDB;Password=Prodrtm@123;");
        SqlConnection devCon = new SqlConnection(@"Data Source=10.55.5.40,1433;Initial Catalog=Real_Time_Metrics_Dev;User ID=PRODRTMDB;Password=Prodrtm@123;");
        DataTable dt = new DataTable();
        int flagSKU = 0;
        int flagSKUQC = 0;
        int flagCMPQC = 0;
        int flagPSL = 0;
        public Form1()
        {
            InitializeComponent();
        }
        private static TimeZoneInfo INDIAN_ZONE = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");
        private void Form1_Load(object sender, EventArgs e)
        {
            //GetPSL_QC();
         
            timerSKU.Enabled = true;
        }

        //IPV-Invoices
        private void GetData()
        {
            try
            {
                dt = new DataTable();
                WebReference.Service1 objSku = new WebReference.Service1();
                objSku.Timeout = int.MaxValue;
                dt = objSku.GetIPSKU();

                if (dt.Rows.Count > 0)
                {
                    dt.Rows.Cast<DataRow>().ToList().ForEach(r => r.SetField("VALIDATED_DATETIME", UpdateTime(Convert.ToDateTime(r["VALIDATED_DATETIME"]))));
                    openConnection();
                    SqlBulkCopy objbulk = new SqlBulkCopy(con);
                    //assigning Destination table name  
                    objbulk.DestinationTableName = "RTM_Sku";
                    
          
                    //Mapping Table column  
                    objbulk.ColumnMappings.Add("PROCESSOR", "PROCESSOR");
                    objbulk.ColumnMappings.Add("fullname", "fullname");
                    objbulk.ColumnMappings.Add("VALIDATED_BY", "VALIDATED_BY");
                    objbulk.ColumnMappings.Add("SKU_NUMBER", "SKU_NUMBER");
                    objbulk.ColumnMappings.Add("DATE_ADDED", "DATE_ADDED");
                    objbulk.ColumnMappings.Add("CLIENT", "CLIENT");
                    objbulk.ColumnMappings.Add("TSHEETS_CLIENT_CODE", "TSHEETS_CLIENT_CODE");
                    objbulk.ColumnMappings.Add("DATE_FINISHED", "DATE_FINISHED");
                    objbulk.ColumnMappings.Add("VALIDATED_DATETIME", "VALIDATED_DATETIME");
                    objbulk.ColumnMappings.Add("EDI_VAN", "EDI_VAN");
                    //inserting bulk Records into DataBase   
                    objbulk.WriteToServer(dt);

                    con.Close();
                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show(ex.Message);    
                WriteToErrorLog(ex.Message, ex.StackTrace, "SKU");
            }
            
        }
        //CMP Invoices
        private void GetCMPData()
        {
            try
            {
                dt = new DataTable();
                WebReference.Service1 objCMP = new WebReference.Service1();
                dt = objCMP.GetCMP();
                if (dt.Rows.Count > 0)
                {
                    //dt.Rows.Cast<DataRow>().ToList().ForEach(r => r.SetField("validatedDateBackend", UpdateTime(Convert.ToDateTime(r["validatedDateBackend"]))));
                    //openDevConnection();
                    openConnection();
                    SqlBulkCopy objbulk = new SqlBulkCopy(con);
                    //assigning Destination table name  
                    objbulk.DestinationTableName = "RTM_CMP";

                    //Mapping Table column  
                    objbulk.ColumnMappings.Add("invoiceSubId", "invoiceSubId");
                    objbulk.ColumnMappings.Add("assignedTo", "assignedTo");
                    objbulk.ColumnMappings.Add("validatedBy", "validatedBy");
                    objbulk.ColumnMappings.Add("batchDate", "batchDate");
                    objbulk.ColumnMappings.Add("custAbbr", "custAbbr");
                    objbulk.ColumnMappings.Add("customer", "customer");
                    objbulk.ColumnMappings.Add("buildDateBackend", "buildDateBackend");                    
                    objbulk.ColumnMappings.Add("validatedDateBackend", "validatedDateBackend");
                    
                    //inserting bulk Records into DataBase   
                    objbulk.WriteToServer(dt);
                    con.Close();
                    //devCon.Close();
                }
            }
            catch (Exception ex)
            {
                WriteToErrorLog(ex.Message, ex.StackTrace, "CMP");
            }
        }
        //IPV -QC
        private void GetSKUQC()
        {
            try
            {
                dt = new DataTable();
                WebReference.Service1 objSku = new WebReference.Service1();
                objSku.Timeout = int.MaxValue;
                dt = objSku.GetQCSKU();

                if (dt.Rows.Count > 0)
                {
                    dt.Rows.Cast<DataRow>().ToList().ForEach(r => r.SetField("DATE_FINISHED", UpdateTime(Convert.ToDateTime(r["DATE_FINISHED"]))));
                    openConnection();
                    SqlBulkCopy objbulk = new SqlBulkCopy(con);
                    //assigning Destination table name  
                    objbulk.DestinationTableName = "RTM_Sku_QC";
                    
                    //Mapping Table column  
                    objbulk.ColumnMappings.Add("PROCESSOR", "PROCESSOR");
                    objbulk.ColumnMappings.Add("fullname", "fullname");
                    objbulk.ColumnMappings.Add("VALIDATED_BY", "VALIDATED_BY");
                    objbulk.ColumnMappings.Add("SKU_NUMBER", "SKU_NUMBER");
                    objbulk.ColumnMappings.Add("DATE_ADDED", "DATE_ADDED");
                    objbulk.ColumnMappings.Add("CLIENT", "CLIENT");
                    objbulk.ColumnMappings.Add("TSHEETS_CLIENT_CODE", "TSHEETS_CLIENT_CODE");
                    objbulk.ColumnMappings.Add("DATE_FINISHED", "DATE_FINISHED");
                    objbulk.ColumnMappings.Add("VALIDATED_DATETIME", "VALIDATED_DATETIME");
                    objbulk.ColumnMappings.Add("EDI_VAN", "EDI_VAN");
                    //inserting bulk Records into DataBase   
                    objbulk.WriteToServer(dt);
                    con.Close();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                WriteToErrorLog(ex.Message, ex.StackTrace, "SKU_QC");
            }
        }

        //IPV Report
        //public DataTable GetSKUQC()
        //{
        //    dt = new DataTable();
        //    string query = "select PROCESSOR, fullname, VALIDATED_BY, SKU_NUMBER, DATE_ADDED, CLIENT, TSHEETS_CLIENT_CODE, DATE_FINISHED, VALIDATED_DATETIME, EDI_VAN from sku left join users on VALIDATED_BY =  username left join client on CLIENT = CLIENT_NAME where users.group_num >= 50 and users.ACTIVE =1 and VALIDATED_DATETIME BETWEEN CURDATE() - INTERVAL 1 DAY AND CURDATE() order by VALIDATED_DATETIME";

        //    if (this.OpenConnection() == true)
        //    {
        //        MySqlCommand cmd = new MySqlCommand(query, connection);
        //        cmd.CommandTimeout = int.MaxValue;

        //        using (MySqlDataAdapter da = new MySqlDataAdapter(cmd))
        //        {
        //            da.Fill(dt);
        //        }
        //        this.CloseConnection();
        //    }

        //    return dt;
        //}
        

        private void GetPSL_QC()
        {
            try
            {
                dt = new DataTable();
                //disbling Service***************
                WebReference.Service1 objSku = new WebReference.Service1();
                dt = objSku.GetQCPSL();

                if (dt.Rows.Count > 0)
                {
                    openConnection();
                    SqlBulkCopy objbulk = new SqlBulkCopy(con);
                    objbulk.DestinationTableName = "RTM_PSL_QC";
                    objbulk.ColumnMappings.Add("PKEY", "PKEY");
                    objbulk.ColumnMappings.Add("IPN", "IPN");
                    objbulk.ColumnMappings.Add("client", "client");
                    objbulk.ColumnMappings.Add("TSHEETS_CLIENT_CODE", "TSHEETS_CLIENT_CODE");
                    objbulk.ColumnMappings.Add("validatedBy", "validatedBy");
                    objbulk.ColumnMappings.Add("fullname", "fullname");
                    objbulk.ColumnMappings.Add("validatedDateTime", "validatedDateTime");

                    objbulk.WriteToServer(dt);

                    con.Close();
                }
            }
            catch (Exception ex)
            {                
                WriteToErrorLog(ex.Message, ex.StackTrace, "PSL_QC");
            }
        }
        //CMP-QC
        private void GetCMPQC()
        {
            try
            {
                dt = new DataTable();
                WebReference.Service1 objCMP = new WebReference.Service1();
                dt = objCMP.GetQCCMP();
                if (dt.Rows.Count > 0)
                {
                    //dt.Rows.Cast<DataRow>().ToList().ForEach(r => r.SetField("validatedDateBackend", UpdateTime(Convert.ToDateTime(r["validatedDateBackend"]))));
                    //openDevConnection();
                    openConnection();
                    SqlBulkCopy objbulk = new SqlBulkCopy(con);
                    //assigning Destination table name  
                    objbulk.DestinationTableName = "RTM_CMP_QC";

                    //Mapping Table column  
                    objbulk.ColumnMappings.Add("invoiceSubId", "invoiceSubId");
                    objbulk.ColumnMappings.Add("assignedTo", "assignedTo");
                    objbulk.ColumnMappings.Add("validatedBy", "validatedBy");
                    objbulk.ColumnMappings.Add("batchDate", "batchDate");
                    objbulk.ColumnMappings.Add("custAbbr", "custAbbr");
                    objbulk.ColumnMappings.Add("customer", "customer");
                    objbulk.ColumnMappings.Add("buildDateBackend", "buildDateBackend");
                    objbulk.ColumnMappings.Add("validatedDateBackend", "validatedDateBackend");

                    //inserting bulk Records into DataBase   
                    objbulk.WriteToServer(dt);
                    con.Close();
                    //devCon.Close();
                }
            }
            catch (Exception ex)
            {
                WriteToErrorLog(ex.Message, ex.StackTrace, "CMPQC");
            }
        }

        public static DateTime UpdateTime(DateTime defaultTime)
        {
            DateTime dtTime = DateTime.MinValue;
            
            
            if (defaultTime < (DateTime)SqlDateTime.MinValue)
            {
                dtTime = Convert.ToDateTime("1/1/1753 12:00:00 AM");
            }
            else
            {
                dtTime = defaultTime;
            }

            return dtTime;
        }

        public void openConnection()
        {
            //Stoting connection string   
            if (con.State == ConnectionState.Open)
            {
                con.Close();
            }

            con.Open();

        }

        public void openDevConnection()
        {
            //Stoting connection string   
            if (devCon.State == ConnectionState.Open)
            {
                devCon.Close();
            }

            devCon.Open();

        }

        private void timerSKU_Tick(object sender, EventArgs e)
        {
            try
            {
                DateTime indianTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, INDIAN_ZONE);

                if (indianTime.Hour == 3)
                {
                    if (flagSKU == 0)
                    {
                        flagSKU = 1;
                        GetData();                       
                    }
                }

                if (indianTime.Hour == 4)
                {
                    if (flagSKUQC == 0)
                    {
                        flagSKUQC = 1;
                        GetSKUQC();
                    }
                }

                if (indianTime.Hour == 5)
                {
                    if (flagCMPQC == 0)
                    {
                        flagCMPQC = 1;
                        GetCMPData();
                        GetCMPQC();
                    }
                }

                if (indianTime.Hour == 6)
                {
                    if (flagPSL == 0)
                    {
                        flagPSL = 1;
                        GetPSL_QC();
                    }
                }

                if (indianTime.Hour == 7)
                {
                    flagSKU = 0;
                    flagSKUQC = 0;
                    flagCMPQC = 0;
                    flagPSL = 0;
                }
            }
            catch (Exception)
            {
                
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
    }
}
