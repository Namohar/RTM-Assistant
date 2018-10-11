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
using System.Data.OleDb;


namespace SplitMeetingDuration
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string query;
        SqlConnection con = new SqlConnection(@"Data Source=10.55.5.40,1433;Initial Catalog=RTM_Global_Test;User ID=PRODRTMDB;Password=Prodrtm@123;");
        DataTable dtClients = new DataTable();
        DataTable dtMeetings = new DataTable();
        DataTable dtSelClients = new DataTable();
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                dtMeetings = ConvertExcelToDataTable(@"D:\RTM\meetings.xlsx");
                dtClients = GetClietData("07/01/2017", "07/21/2017");

                if (dtMeetings.Rows.Count > 0)
                {
                    foreach (DataRow drMeet in dtMeetings.Rows)
                    {
                        string LAID = drMeet["LA_ID"].ToString();
                        string user = drMeet["LA_User_Name"].ToString();
                        string date = Convert.ToDateTime(drMeet["LA_TimeDate"]).ToShortDateString();
                        string duration = drMeet["LA_Duration"].ToString();
                        string expression = "R_User_Name='" + user + "' and R_TimeDate= '" + date + "'";

                        var rows = dtClients.Select(expression);

                        if (rows.Any())
                        {
                            dtSelClients = new DataTable();
                            dtSelClients = rows.CopyToDataTable();

                            double splitTime = (Convert.ToDouble(duration) / dtSelClients.Rows.Count);

                            dtSelClients.Rows.Cast<DataRow>().ToList().ForEach(r => r.SetField("R_Duration", UpdateTime(r["R_Duration"].ToString(), splitTime)));

                            if (dtSelClients.Rows.Count > 0)
                            {
                                foreach (DataRow finalRow in dtSelClients.Rows)
                                {
                                    UpdateClientDuration(finalRow["R_ID"].ToString(), finalRow["R_Duration"].ToString(), LAID);
                                }

                                DeleteLog(LAID);
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                
                throw;
            }
           
        }

        public static String UpdateTime(string duraton, double splitTime)
        {
            Double clientTime = 0;

            string totalDuration = "00:00:00";

            if (string.IsNullOrEmpty(duraton))
            {
                duraton = "00:00:00";
            }
            Double hrs = TimeSpan.Parse(duraton).TotalHours;

            clientTime = hrs + splitTime;


            var hours = clientTime.ToString().Split('.')[0];
            var minutes = ((clientTime * 60) % 60).ToString().Split('.')[0];
            var seconds = ((clientTime * 3600) % 60).ToString().Split('.')[0];


            totalDuration = hours + ":" + minutes + ":" + seconds;

            return totalDuration;
        }

        private DataTable GetClietData(string from, string to)
        {
            try
            {
                query = "select R_ID, R_User_Name, CL_ClientName, R_Duration, CONVERT(VARCHAR(12), R_TimeDate, 101) as R_TimeDate from RTM_Records left join RTM_Client_List on R_Client = CL_ID left join RTM_Team_List on R_TeamId = T_ID where CL_ClientName !='Internal' and CL_ClientName !='Jury Duty' and CL_ClientName !='Personal/Sick Time' and CL_ClientName !='Public Holiday' and CL_ClientName !='Vacation' and CL_ClientName !='Bereavement' and CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, R_TimeDate))) Between '" + from + "' and '" + to + "' and R_Duration != 'HH:MM:SS' and T_Location ='IND'";
                using (SqlDataAdapter da = new SqlDataAdapter(query, con))
                {
                    da.Fill(dtClients);
                }
            }
            catch (Exception)
            {
                
                throw;
            }
            

            return dtClients;
        }
      
        public static DataTable ConvertExcelToDataTable(string FileName)  
        {  
            DataTable dtResult = null;  
            int totalSheet = 0; //No of sheets on excel file  
            using(OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"))  
            {  
                objConn.Open();  
                OleDbCommand cmd = new OleDbCommand();  
                OleDbDataAdapter oleda = new OleDbDataAdapter();  
                DataSet ds = new DataSet();  
                DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);  
                string sheetName = string.Empty;  
                if (dt != null)  
                {  
                    var tempDataTable = (from dataRow in dt.AsEnumerable()  
                    where!dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")  
                    select dataRow).CopyToDataTable();  
                    dt = tempDataTable;  
                    totalSheet = dt.Rows.Count;  
                    sheetName = dt.Rows[0]["TABLE_NAME"].ToString();  
                }  
                cmd.Connection = objConn;  
                cmd.CommandType = CommandType.Text;  
                cmd.CommandText = "SELECT * FROM [" + sheetName + "]";  
                oleda = new OleDbDataAdapter(cmd);  
                oleda.Fill(ds, "excelData");  
                dtResult = ds.Tables["excelData"];  
                objConn.Close();  
                return dtResult; //Returning Dattable  
            }  
        }

        private void UpdateClientDuration(string id, string duration, string LAID)
        {
            int result;
            using (SqlCommand cmd = new SqlCommand("update RTM_Records set R_Duration ='"+ duration +"' where R_ID ='"+ id +"'", con))
            {
                cmd.CommandTimeout = int.MaxValue;
                con.Open();
                result = cmd.ExecuteNonQuery();
                con.Close();
            }
           
        }

        private void DeleteLog(string LAID)
        {
            using (SqlCommand cmd = new SqlCommand("Delete from RTM_Log_Actions where LA_ID ='" + LAID + "'", con))
            {
                cmd.CommandTimeout = int.MaxValue;
                con.Open();
                int result = cmd.ExecuteNonQuery();
                con.Close();
            }
        }
    }
}
