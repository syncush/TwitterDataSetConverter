using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ParseJson
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {   // Open the text file using a stream reader.
                var csv = new StringBuilder();
                Microsoft.Office.Interop.Excel.Application oXL;
                Microsoft.Office.Interop.Excel._Workbook oWB;
                Microsoft.Office.Interop.Excel._Worksheet oSheet;
                Microsoft.Office.Interop.Excel.Range oRng;
                object misValue = System.Reflection.Missing.Value;
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;
                oWB = (Microsoft.Office.Interop.Excel.Workbook)(oXL.Workbooks.Add(""));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                JObject jobj = null;
                using (StreamReader sr = new StreamReader(@"C:\Users\sync\Desktop\ParseJSON\ParseJson\ParseJson\koko.json"))
                {
                    for (int i = 1; i <= 15000; ++i)
                    {
                        string singleLine = sr.ReadLine();
                         jobj = JObject.Parse(singleLine);
                        string tweetText = jobj["text"].Value<string>().Replace(Environment.NewLine,"");
                        string createdAt = jobj["created_at"].Value<string>();
                        string timeZone = (jobj["user"])["time_zone"].Value<string>();
                        int followCount = (jobj["user"])["followers_count"].Value<int>();
                        int friendCount = (jobj["user"])["friends_count"].Value<int>();
                        oSheet.Cells[1][i] = tweetText;
                        oSheet.Cells[2][i] = createdAt;
                        oSheet.Cells[3][i] = followCount;
                        oSheet.Cells[4][i] = friendCount;
                        oSheet.Cells[5][i] = timeZone;
                        //csv.AppendLine(string.Format("{0}%%{1}\n", m.Groups["yolo"].Value, "1"));

                    }
                    oXL.Visible = false;
                    oXL.UserControl = false;
                    oWB.SaveAs(@"C:\Users\sync\Desktop\ParseJSON\ParseJson\dataset.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    oWB.Close();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(e.Message);
            }
        }
    }
}
