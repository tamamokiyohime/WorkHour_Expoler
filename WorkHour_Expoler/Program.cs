using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Data;
using System.Text.RegularExpressions;
//關閉部分針對null的警示(.net 6)
#pragma warning disable CS8604
#pragma warning disable CS8602
#pragma warning disable CS8629

namespace Working_hours_Exporter
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/calendar-dotnet-quickstart.json
        static string[] Scopes = { CalendarService.Scope.CalendarEventsReadonly };
        static string ApplicationName = "Google Calendar API .NET Quickstart";
        static string[] weekday = { "一", "二", "三", "四", "五", "六", "日" };
        static string calenderID = "8k13f6uukhu1r12no3p0hbfin8@group.calendar.google.com";
        static void Main(string[] args)
        {
            Console.Title = "Working hour Exporter -> 工讀值班班表(集思軒)";
            // Create Google Calendar API service.
            UserCredential credential;
            Console.WriteLine("Google Login...");
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            } 
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            ////取得登入人的日曆清單
            //CalendarListResource.ListRequest request = service.CalendarList.List();
            //CalendarList cal = request.Execute();
            //Console.Write("Auto catching calender ID of 工讀值班班表(集思軒)..... ->");
            //bool getCalender = false;
            //CalendarListEntry calender = new CalendarListEntry();
            ////尋找日曆清單裡是否有「工讀值班班表(集思軒)」，並取得calenderID
            //foreach (var summary in cal.Items)
            //{
            //    if (string.Equals(summary.Summary, "工讀值班班表(集思軒)"))
            //    {
            //        calender = summary;
            //        getCalender = true;
            //    }
            //}
            //if (getCalender)
            //    Console.WriteLine(calender.Id.ToString());
            //else
            //{
            //    Console.WriteLine("Unable to get calender ID.... exit");
            //    return;
            //}
            Console.WriteLine("----------------------------");
            Console.Write("請輸入下載日期(yyyy-MM)：");
            int inputyear = 0;
            int inputmonth = 0;
            //檢查輸入的日期格式是否正確
            while (true)
            {
                int rc = CheckInputDateFormate(Console.ReadLine(), out inputyear, out inputmonth);
                if (rc == 0) break;
                else Console.Write(CheckInputDateErrorCode(rc));
            }
            bool split4hour = true;
            Console.Write("4小時自動拆分功能 0:關閉/1:開啟 [1]：");
            CheckInput4hoursplit(Console.ReadLine(), out split4hour);
            
            //查詢calenderID中指定日期範圍的全部事件
            //EventsResource.ListRequest request1 = service.Events.List(calender.Id);
            EventsResource.ListRequest request1 = service.Events.List(calenderID);
            //Google apis中指定TimeMin與TimeMax的格式必須是RCF1123，不過這裡用Datetime傳所以不照RCF寫
            request1.TimeMin = new DateTime(inputyear, inputmonth, 01, 01, 00, 00, DateTimeKind.Local);
            request1.TimeMax = new DateTime(inputyear, inputmonth, DateTime.DaysInMonth(inputyear, inputmonth), 23, 59, 59, DateTimeKind.Local);request1.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;
            request1.ShowDeleted = false;
            request1.SingleEvents = true;
            Events events = request1.Execute();

            //建立資料集並依照標題建立資料表
            DataSet dataSet = new DataSet();
            if (events.Items != null && events.Items.Count > 0)
            {
                foreach (var eventItem in events.Items)
                {
                    
                    if (eventItem.Summary != null && eventItem.Start.DateTime != null)
                    {
                        //建立資料表
                        if (!dataSet.Tables.Contains(eventItem.Summary))
                        {
                            dataSet.Tables.Add(eventItem.Summary);
                            dataSet.Tables[eventItem.Summary].Columns.Add("Date");
                            dataSet.Tables[eventItem.Summary].Columns.Add("Weekday");
                            dataSet.Tables[eventItem.Summary].Columns.Add("Time");
                            dataSet.Tables[eventItem.Summary].Columns.Add("Duration");
                        }
                        //將事件填入資料表中
                        DateTime start = (DateTime)eventItem.Start.DateTime;
                        DateTime end = (DateTime)eventItem.End.DateTime;
                        TimeSpan duration = end - start;
                        

                        if (split4hour)
                        {
                            TimeSpan restime = new(0, 30, 0);
                            TimeSpan maxtime = new(4, 0, 0);
                            DateTime nextstart = start;

                            while (true)//自動將大於4小時的時間拆開，並在中間隔出30分鐘空白
                            {
                                DataRow row = dataSet.Tables[eventItem.Summary].NewRow();
                                row["Date"] = start.ToString("dd");
                                row["Weekday"] = weekday[(int)start.DayOfWeek];
                                TimeSpan addup = duration > maxtime ? new(4, 0, 0) : duration;
                                row["Time"] = nextstart.ToString("HH：mm") + "～" + (nextstart + addup).ToString("HH：mm");
                                row["Duration"] = addup.TotalHours;
                                nextstart += addup + restime;
                                dataSet.Tables[eventItem.Summary].Rows.Add(row);
                                if (duration <= maxtime) break;
                                else    duration -= maxtime;
                            }
                        }
                        else
                        {
                            DataRow row = dataSet.Tables[eventItem.Summary].NewRow();
                            row["Date"] = start.ToString("dd");
                            row["Weekday"] = weekday[(int)start.DayOfWeek];
                            row["Time"] = start.ToString("HH：mm") + "～" + end.ToString("HH：mm");
                            row["Duration"] = duration.TotalHours;
                            dataSet.Tables[eventItem.Summary].Rows.Add(row);
                        }
                    }
                }
            }
            Console.WriteLine("Table count:" + dataSet.Tables.Count);
            //將資料表輸出成docx
            foreach (DataTable table in dataSet.Tables)
            {
                //Console.WriteLine(tabletostring(table));
                Getdocx(table, inputyear.ToString("0000") + "-" + inputmonth.ToString("00"));
            }
            Console.WriteLine("Process Finish. Closing in ");
            for (int i = 3; i > 0; i--) 
            {
                Console.Write(i.ToString() + ".. ");
                Thread.Sleep(1000);
            }
        }
        static int CheckInputDateFormate(string input, out int year, out int day)
        {
            year = 0;
            day = 0;
            if (String.IsNullOrEmpty(input))            return 1;
            if (!Regex.IsMatch(input, @"^\d{4}-\d{2}")) return 2;
            string[] inputvalue = input.Split('-');
            if (!int.TryParse(inputvalue[0], out year) || year < 2000 || year > DateTime.Now.Year) return 3;
            if (!int.TryParse(inputvalue[1], out day) || day < 1 || day > 12) return 4;
            return 0;
        }
        static int CheckInput4hoursplit(string input, out bool active)
        {
            active = true;
            if (!Regex.IsMatch(input, @"^\d{1}"))   return 1;
            if (String.IsNullOrEmpty(input))        return 2;
            if (!bool.TryParse(input, out active))  return 3;
            return 0;
        }
        static string CheckInputDateErrorCode(int i)
        {
            switch (i)
            {
                case 0: return "Pass";
                case 1: return "空字串，重新輸入：";
                case 2: return "字串格式不符，重新輸入：";
                case 3: return "年分格式不符，重新輸入：";
                case 4: return "月分格式不符，重新輸入：";
                default: return "NP";
                
            }
        }
        static bool Getdocx(DataTable dt, string folder)
        {
            //將模板檔讀到記憶流中
            byte[] byteArray = File.ReadAllBytes("template.docx");
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, byteArray.Length);
                //以記憶流中的模板檔進行編輯
                using (WordprocessingDocument doc = WordprocessingDocument.Open(stream, true))
                {
                    // Create an empty table.
                    Table table = new Table();
                    // Create a TableProperties object and specify its border information.
                    table.AppendChild(
                        new TableProperties(
                            new TableBorders(
                                new TopBorder()                 { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                                new BottomBorder()              { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                                new LeftBorder()                { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                                new RightBorder()               { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                                new InsideHorizontalBorder()    { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                                new InsideVerticalBorder()      { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 }
                                )
                            )
                        );
                    //寫入行首
                    table.Append(GetRow("日期", "星期", "工作項目", "工作地點", "起迄時間", "時數", "工讀生簽章", "館員核對", true));
                    float totalHour = 0;
                    int rowcount = 1;
                    foreach (DataRow row in dt.Rows)
                    {
                        if (rowcount == 23) //到頁尾的時候加入下頁行以及在下一頁寫入行首
                        {
                            table.Append(GetMergedRow("接　　續　　下　　頁", true));
                            table.Append(GetRow("日期", "星期", "工作項目", "工作地點", "起訖時間", "時數", "工讀生簽章", "館員核對", true));
                            rowcount = 1;
                        }
                        totalHour += float.Parse(row["Duration"].ToString());   //計算總時數
                        table.Append(GetRow(row["Date"].ToString(), row["Weekday"].ToString(), "文件整理、資料建檔、電腦測試", "集思軒", row["Time"].ToString(), row["Duration"].ToString(), "", ""));//寫入資料行
                        rowcount++;
                    }

                    for(int i = rowcount; i < 23; i++)  //將表格補到滿頁
                    {
                        table.Append(GetRow("", "", "", "", "", "", "", ""));
                    }
                    //寫入尾行
                    table.Append(GetMergedRow("承辦人：　　　　　　　　　　組長（核定）：　　　　　　　　全月工作時數：" + totalHour.ToString("#00"), false));
                    // Append the table to the document.
                    //doc.MainDocumentPart.Document.Body.Append(table);
                    //將表格寫入到doc中
                    doc.MainDocumentPart.Document.Body.AddChild(table);                    

                    // Save changes to the MainDocumentPart.
                    doc.MainDocumentPart.Document.Save();
                }
                // Save the file with the new name
                if (!Directory.Exists("Output\\" + folder)) Directory.CreateDirectory("Output\\" + folder);             //檢查目標資料夾是否存在
                File.WriteAllBytes("Output\\" + folder + "\\" + dt.TableName.ToString() + ".docx", stream.ToArray());   //將記憶流資料寫到硬碟中
            }
            return true;
        }
        static TableRow GetRow(string date, string weekday, string content, string place, string time, string hour, string sign, string check, bool title = false)
        {
            TableRow tr = new TableRow();
            tr.Append(GetCell(6, date, title));
            tr.Append(GetCell(6, weekday, title));
            tr.Append(GetCell(28, content, title));
            tr.Append(GetCell(8, place, title));
            tr.Append(GetCell(18, time, title));
            tr.Append(GetCell(6, hour, title));
            tr.Append(GetCell(14, sign, title));
            tr.Append(GetCell(14, check, title));
            return tr;
        }
        static TableRow GetMergedRow(string text, bool iscenter)
        {
            TableRow tr = new TableRow();
            tr.Append(GetMergedCell(true, iscenter, text));
            tr.Append(GetMergedCell(false));
            tr.Append(GetMergedCell(false));
            tr.Append(GetMergedCell(false));
            tr.Append(GetMergedCell(false));
            tr.Append(GetMergedCell(false));
            tr.Append(GetMergedCell(false));
            tr.Append(GetMergedCell(false));
            return tr;
        }
        static TableCell GetCell(int percent, string text, bool isTitle = false)
        {
            TableCell tc = new TableCell();
            tc.Append(
                new TableCellProperties(
                    new TableCellWidth()
                    {
                        Type = TableWidthUnitValues.Pct,
                        Width = percent.ToString()
                    },
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                    )
                );
            tc.Append(
                new Paragraph(
                    new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Center }
                        ),
                    new Run(
                        new RunProperties(new Bold() { Val = OnOffValue.FromBoolean(isTitle) }),
                        new Text(text)
                        )
                    )
                );
            return tc;
        }
        static TableCell GetMergedCell(bool isStart, bool isCenter = false, string text = "")
        {
            TableCell tc = new TableCell();
            tc.Append(
                new TableCellProperties(
                    new HorizontalMerge() { Val = isStart ? MergedCellValues.Restart : MergedCellValues.Continue }
                    )
                );
            // Specify the table cell content.
            tc.Append(
                new Paragraph(
                    new ParagraphProperties(
                        new Justification() { Val = isCenter ? JustificationValues.Center : JustificationValues.Left }
                        ),
                    new Run(
                        new RunProperties(new Bold() { Val = OnOffValue.FromBoolean(true) }),
                        new Text(isStart ? text : "")
                        )
                    )
                );
            return tc;
        }
        static string tabletostring(DataTable dt)
        {
            string s = dt.TableName + "\n";
            foreach (DataColumn col in dt.Columns)
            {
                s += col.ColumnName;
                s += "\t";
            }
            s += "\n";
            foreach (DataRow row in dt.Rows)
            {
                foreach (var data in row.ItemArray)
                {
                    s += data.ToString();
                    s += "\t";
                }
                s += "\n";
            }
            return s;
        }
    }

}