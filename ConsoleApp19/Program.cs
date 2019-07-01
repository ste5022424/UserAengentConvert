using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using UAParser;

namespace ConsoleApp19
{
    internal class Program
    {
        public static string getus(string uaString)
        {
            // get a parser with the embedded regex patterns
            var uaParser = Parser.GetDefault();

            ClientInfo c = uaParser.Parse(uaString);

            //Console.WriteLine(c.UserAgent.Family); // => "Mobile Safari"
            //Console.WriteLine(c.UserAgent.Major);  // => "5"
            //Console.WriteLine(c.UserAgent.Minor);  // => "1"

            //Console.WriteLine(c.OS.Family);        // => "iOS"
            //Console.WriteLine(c.OS.Major);         // => "5"
            //Console.WriteLine(c.OS.Minor);         // => "1"

            //Console.WriteLine(c.Device.Family);    // => "iPhone"

            return c.OS.Family;
        }

        private static void Main(string[] args)
        {
            try
            {
                Console.WriteLine($"==程式開始==");
                string exportpath = $"D:\\APINew{System.DateTime.Now.ToString("yyyyMMddHHmmss")}";
                Console.WriteLine($"匯出Excel路徑 : {exportpath}");
                var fileName = string.Format("{0}\\APIData.xlsx", Directory.GetCurrentDirectory());
                Console.WriteLine($"讀取Excel路徑 : {fileName}");

                string strConn = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={fileName};Extended Properties =\"Excel 12.0 Xml;HDR=YES\";";
                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                string strExcel = "";
                OleDbDataAdapter myCommand = null;
                DataSet ds = null;
                strExcel = "select * from [APIData$]";
                myCommand = new OleDbDataAdapter(strExcel, strConn);
                ds = new DataSet();
                myCommand.Fill(ds, "APIData");

                Console.WriteLine($"資料轉換開始，共 {ds.Tables[0].Rows.Count} 筆資料");

                var UserAgentDataList = ds.Tables[0].AsEnumerable().Select(x => new my
                {
                    //  Date = x["APIData"].ToString(),
                    Host = x["HttpHost"].ToString(),
                    URL = x["HttpURL"].ToString(),
                    PlatForm = getus(x["UserAgent"].ToString()),
                    Count = x["Count"].ToString(),
                }).ToList();

                var GroupByData = UserAgentDataList.GroupBy(item => new
                {
                    item.PlatForm,
                    item.URL,
                    //item.Date,
                    item.Host
                })
                .Select(group => new my
                {
                    // Date = group.Key.Date,
                    Host = group.Key.Host,
                    URL = group.Key.URL,
                    PlatForm = group.Key.PlatForm,
                    Count = group.Select(y => y.Count).Sum(a => Convert.ToInt64(a)).ToString(),
                }
                )
                .OrderBy(x => x.Date)
                .ToList();

                Console.WriteLine($"資料轉換 完成，共 {GroupByData.Count} 筆資料");
                Console.WriteLine("匯出Excel");
                SaveDataToExcelFile(GroupByData, exportpath);
                Console.WriteLine($"匯出Excel 成功， 路徑 {exportpath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("發生錯誤" + ex);
            }
        }

        private static void SaveDataToExcelFile(List<my> studentList, string filePath)
        {
            object misValue = System.Reflection.Missing.Value;
            Application xlApp = new Application();
            Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
            Worksheet xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            PropertyInfo[] props = GetPropertyInfoArray();
            for (int i = 0; i < props.Length; i++)
            {
                xlWorkSheet.Cells[1, i + 1] = props[i].Name; //write the column name
            }
            for (int i = 0; i < studentList.Count; i++)
            {
                xlWorkSheet.Cells[i + 2, 1] = studentList[i].Date;
                xlWorkSheet.Cells[i + 2, 2] = studentList[i].Host;
                xlWorkSheet.Cells[i + 2, 3] = studentList[i].URL;
                xlWorkSheet.Cells[i + 2, 4] = studentList[i].PlatForm;
                xlWorkSheet.Cells[i + 2, 5] = studentList[i].Count;
            }
            try
            {
                xlWorkBook.SaveAs(filePath, XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine("發生錯誤" + ex);
            }
        }

        private static PropertyInfo[] GetPropertyInfoArray()
        {
            PropertyInfo[] props = null;
            try
            {
                Type type = typeof(my);
                object obj = Activator.CreateInstance(type);
                props = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            }
            catch (Exception ex)
            {
                Console.WriteLine("發生錯誤" + ex);
            }
            return props;
        }

        public class my
        {
            public string Date { get; set; }
            public string Host { get; set; }
            public string URL { get; set; }
            public string PlatForm { get; set; }
            public string Count { get; set; }
        }
    }
}