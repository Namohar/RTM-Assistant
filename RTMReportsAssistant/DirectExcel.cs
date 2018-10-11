using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace RTMReportsAssistant
{
    public class DirectExcel : System.Web.UI.WebControls.GridView
    {
        public System.IO.MemoryStream ExportToStream(DataTable objDt)
        {

            //remove any columns specified.

            //foreach (string colName in ColumnsToRemove)
            //{

            //    objDt.Columns.Remove(colName);

            //}

            this.DataSource = objDt;

            this.DataBind();



            System.IO.StringWriter sw = new System.IO.StringWriter();

            System.Web.UI.HtmlTextWriter hw = new System.Web.UI.HtmlTextWriter(sw);

            this.RenderControl(hw);

            string content = sw.ToString();

            byte[] byteData = Encoding.Default.GetBytes(content);


            System.IO.MemoryStream mem = new System.IO.MemoryStream();

            mem.Write(byteData, 0, byteData.Length);

            mem.Flush();

            mem.Position = 0; //reset position to the begining of the stream

            return mem;
        } 
    }
}
