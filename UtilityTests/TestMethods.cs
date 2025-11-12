using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using OfficeOpenXml;

namespace UtilityTests
{
    [TestClass]
    public class TestTemp
    {
        [TestMethod]
        public void TestMethod1()
        {
            try
            {
                string filePath = @"D:\TestFiles\ecalendarTableData.txt";
                string fileContent = File.ReadAllText(filePath);
                List<CalendaModel> files = JsonConvert.DeserializeObject<List<CalendaModel>>(fileContent);

                List<OutModel> result = files.Select(ToOutModel).OrderBy(m => m.Date).ToList();
                string filePathOut = @"D:\TestFiles\OutData.txt";
                JsonSerializerSettings settings = new JsonSerializerSettings
                {
                    Culture = new CultureInfo("zh-CN"),
                    DateFormatString = "yyyy年MM月dd日 HH:mm:ss dddd",
                    NullValueHandling = NullValueHandling.Ignore,
                };
                File.WriteAllText(filePathOut, JsonConvert.SerializeObject(result, Formatting.Indented, settings));

                string csvFilePath = @"D:\TestFiles\万年历数据.csv";
                StringBuilder sb = new StringBuilder();
                List<string> titles = new List<string>() { "日期", "类型", "内容", };
                sb.AppendLine(string.Join(",", titles));
                foreach (OutModel item in result)
                {
                    sb.AppendLine($"{item.Date:yyyy年MM月dd日 HH:mm:ss dddd},{(item.Type == 2 ? "生日" : "记事")},{item.Note}");
                }
                File.WriteAllText(csvFilePath, sb.ToString(), Encoding.UTF8);

                CvsToExcel(csvFilePath);
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex);
            }
        }

        public static readonly DateTime Epoch = new DateTime(1970, 1, 1);
        private OutModel ToOutModel(CalendaModel model)
        {
            OutModel result = new OutModel();
            result.Type = model.lineType;
            result.Date = Epoch.AddMilliseconds(model.time).ToLocalTime();
            //result.UpdateDate = Epoch.AddMilliseconds(model.update_time).ToLocalTime();

            List<string> list = new List<string>();
            string title = RemoveHtmlTags(model.title);
            string note = RemoveHtmlTags(model.note);
            list.Add(note);
            list.Add(title);
            list = list.OrderBy(i => i.Length).ToList();

            string content = string.Join(": ", list.Where(i => !string.IsNullOrEmpty(i)))
                                   .Replace("\r", " ")
                                   .Replace("\f", " ")
                                   .Replace("\n", " ")
                                   .Replace("\t", " ")
                                   .Replace("\v", " ");
            if (result.Type == 2)
            {
                result.Birthday = content;
            }
            else
            {
                result.Note = content;
            }
            
            return result;
        }

        static string RemoveHtmlTags(string input)
        {
            return Regex.Replace(input, "<.*?>", string.Empty).Trim();
        }

        public static void CvsToExcel(string csvFilePath)
        {
            string excelFilePath = csvFilePath.Replace(".csv", ".xlsx");
            
            string[] csvLines = File.ReadAllLines(csvFilePath, Encoding.UTF8);
            
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
                
                for (int i = 0; i < csvLines.Length; i++)
                {
                    string[] csvValues = csvLines[i].Split(',');

                    for (int j = 0; j < csvValues.Length; j++)
                    {
                        worksheet.Cells[i + 1, j + 1].Value = csvValues[j];
                    }
                }
                
                worksheet.Cells.AutoFitColumns();
                
                File.WriteAllBytes(excelFilePath, excelPackage.GetAsByteArray());
            }

        }
    }

    #region Test Models

    public class OutModel
    {
        [JsonProperty("日期")]
        public DateTime Date { get; set; }

        //[JsonProperty("更新日期")]
        //public DateTime UpdateDate { get; set; }

        [JsonProperty("记事")]
        public string Note { get; set; }

        [JsonProperty("生日")]
        public string Birthday { get; set; }

        [JsonIgnore]
        public int Type { get; set; }
    }

    public class CalendaModel
    {
        public int id { get; set; }
        public string sid { get; set; }
        public int flag { get; set; }
        public int isSyn { get; set; }
        public long tx { get; set; }
        public int lineType { get; set; }
        public string title { get; set; }
        public string note { get; set; }
        public int catId { get; set; }
        public int isRing { get; set; }
        public string ring { get; set; }
        public int isNormal { get; set; }
        public int syear { get; set; }
        public int smonth { get; set; }
        public int sdate { get; set; }
        public int shour { get; set; }
        public int sminute { get; set; }
        public int nyear { get; set; }
        public int nmonth { get; set; }
        public int ndate { get; set; }
        public int nhour { get; set; }
        public int nminute { get; set; }
        public int advance { get; set; }
        public int cycle { get; set; }
        public int cycleWeek { get; set; }
        public string data { get; set; }
        public string otherData { get; set; }
        public long time { get; set; }
        public int sub_catid { get; set; }
        public long update_time { get; set; }
        public string noteLabelName { get; set; }
    }

    #endregion

}