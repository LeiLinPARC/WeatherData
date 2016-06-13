using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace weatherdatadownload
{
    // change line 61 url, go to www.weatherunderground.com, check the url of the location
    // change line 110, output location
    // change line 126, year, month, day
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public string GetHTML(string url, string encoding)   //get html；
        {
            WebClient web = new WebClient();
            byte[] buffer = web.DownloadData(url);
            return Encoding.GetEncoding(encoding).GetString(buffer);
        }

        public void downloadData(int y, int m, int d) {
            string sd;
            string sm;
            string sy;
            string url;
            string result;
            string dd;
            string ndd;

            sd = Convert.ToString(d);
            sm = Convert.ToString(m);
            sy = Convert.ToString(y);

            DataTable dtDay = new DataTable(sm + "-" + sd + "-" + sy);
            dtDay.Columns.Add("time", System.Type.GetType("System.String"));
            dtDay.Columns.Add("Temp", System.Type.GetType("System.String"));
            dtDay.Columns.Add("Dew Point", System.Type.GetType("System.String"));
            dtDay.Columns.Add("Humidity", System.Type.GetType("System.String"));
            dtDay.Columns.Add("Pressure", System.Type.GetType("System.String"));
            dtDay.Columns.Add("Visibility", System.Type.GetType("System.String"));
            dtDay.Columns.Add("Wind Dir", System.Type.GetType("System.String"));
            dtDay.Columns.Add("Wind Speed", System.Type.GetType("System.String"));
            dtDay.Columns.Add("Gust Speed", System.Type.GetType("System.String"));
            dtDay.Columns.Add("Precip", System.Type.GetType("System.String"));
            dtDay.Columns.Add("Events", System.Type.GetType("System.String"));
            dtDay.Columns.Add("Conditions", System.Type.GetType("System.String"));
            dtDay.Columns.Add("Wind Dir Degree", System.Type.GetType("System.String"));
            dtDay.Columns.Add("DateUTC", System.Type.GetType("System.String"));
            
            url = "https://www.wunderground.com/history/airport/KROC/" + sy + "/" + sm + "/" + sd + "/DailyHistory.html?req_city=Rochester&req_state=NY&req_statename=&reqdb.zip=14580&reqdb.magic=1&reqdb.wmo=99999&format=1";
            result = GetHTML(url, "gb2312");

            Regex check = new Regex("\\n.*<br />");
            MatchCollection marticles = check.Matches(result);
            foreach (Match mar in marticles)
            {
                dd = mar.ToString();

                ndd = dd.Replace("\n", "");
                ndd = ndd.Replace("<br />", "");

                string[] mdd = ndd.Split(',');
                DataRow dr = dtDay.NewRow();
                dr["time"] = mdd[0];
                dr["Temp"] = mdd[1];
                dr["Dew Point"] = mdd[2];
                dr["Humidity"] = mdd[3];
                dr["Pressure"] = mdd[4];
                dr["Visibility"] = mdd[5];
                dr["Wind Dir"] = mdd[6];
                dr["Wind Speed"] = mdd[7];
                dr["Gust Speed"] = mdd[8];
                dr["Precip"] = mdd[9];
                dr["Events"] = mdd[10];
                dr["Conditions"] = mdd[11];
                dr["Wind Dir Degree"] = mdd[12];
                dr["DateUTC"] = mdd[13];
                dtDay.Rows.Add(dr);
            }
            Excel.Application exData = new Excel.Application();
            exData.Workbooks.Add(true);
            int row = 1;
            //for (int i = 0; i < dtDay.Columns.Count; i++)
            //{
            //   exData.Cells[1, i + 1] = dtDay.Columns[i].ColumnName.ToString();
            //}
            for (int i = 0; i < dtDay.Rows.Count; i++)
            {
                for (int j = 0; j < dtDay.Columns.Count; j++)
                {
                    exData.Cells[row, j + 1] = dtDay.Rows[i][j].ToString();
                }
                row++;
            }
            // exData.Visible = true;

            foreach (Excel.Workbook wkb in exData.Workbooks)
            {
                wkb.SaveAs(@"C:\LeiLin\USX28452\Downloads\weatherdata\" + "roc_"+sm + "-" + sd + "-" + sy + ".xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            exData.Quit();
            exData = null;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            int d;
            int m;
            int y;
            d = 1;
            m = 1;
            y = 2012;
            


            for (y = 2012; y <= 2015; y++)// which year to start
            {
                for (m = 1; m <= 12; m++) // which month to start
                {
                    if (m == 1 || m == 3 || m == 5 || m == 7 || m == 8 || m == 10 || m == 12)
                    {
                        for (d = 1; d <= 31; d++)
                        {
                            downloadData(y, m, d);
                        }
                    }
                    else if (m == 2 && (y != 2004 && y != 2008))
                    {
                        for (d = 1; d <= 28; d++)
                        {
                            downloadData(y, m, d);
                        }
                    }
                    else if (m == 2 && (y == 2004 || y == 2008))
                    {
                        for (d = 1; d <= 29; d++)
                        {
                            downloadData(y, m, d);
                        }
                    }
                    else if (m == 4 || m == 6 || m == 9 || m == 11)
                    {
                        for (d = 1; d <= 30; d++)
                        {
                            downloadData(y, m, d);
                        }
                    }
                }
            }
        }
    }
}
