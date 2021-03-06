﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Timers;
using System.Threading;
using System.Xml.XPath;
using Newtonsoft.Json.Linq;

//Telegram API
using Telegram.Bot;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;

// Google Sheet API
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

// HtmlAgilityPack
using HtmlAgilityPack;

// Twitter Korean Processor
using Moda.Korean.TwitterKoreanProcessorCS;

namespace Athena
{
    class CBotClient
    {
        CSystemInfo systemInfo = new CSystemInfo();     // 시스템 정보

        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Clien Delicious Team Notice Bot";
        UserCredential credential;
        SheetsService service;
        CNotice Notice = new CNotice();
        CEasterEgg EasterEgg = new CEasterEgg();
        CUserDirector userDirector = new CUserDirector();
        CNaturalLanguage naturalLanguage = new CNaturalLanguage();
        CNasInfo nasInfo = new CNasInfo();
        CConfig config = new CConfig();

        bool isGoodMorning = false;
        bool isLupinOfWeek = false;

        private Telegram.Bot.TelegramBotClient Bot;// = new Telegram.Bot.TelegramBotClient(strBotToken);

        public void InitBotClient()
        {
            string strPrint = "";
            strPrint += "[ Athena Start ]\n";

            systemInfo.SetStartTime();
            strPrint += "- System Time Loading Completed.\n";

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });


            // json
            config.loadConfig();
#if DEBUG
            Bot = new Telegram.Bot.TelegramBotClient(config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_TEST));
#else
            Bot = new Telegram.Bot.TelegramBotClient(config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_BOT));
#endif

            strPrint += "- External API Loading Completed.\n";


            // 시트에서 유저 정보를 Load
            if (loadUserInfo() == true)
                strPrint += "- User Info Loading Completed.\n";
            else
                strPrint += "- User Info Loading Fail.\n";


            // 백업 파일에서 알림, 메모, 채팅갯수를 Load
            if (loadData() == true)
                strPrint += "- User Data Loading Completed.\n";
            else
                strPrint += "- User Data Loading Fail.\n";


            // NAS 경로 기본값 설정
            nasInfo.CurrentPath = @"D:\CDT\";
            strPrint += "- Administrator Function Loading Completed.\n";

            // 타이머 생성 및 시작
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 5000; // 5초
            timer.Elapsed += new ElapsedEventHandler(timer_Elapsed);
            timer.Start();
            strPrint += "- Thread Create Completed.\n";

            // 아테나 구동 알림            
            strPrint += "- System Time : " + systemInfo.GetNowTime() + "\n";
            strPrint += "- Running Time : " + systemInfo.GetRunningTime() + "\n";
            strPrint += "[ All Completed. ]";

//#if !DEBUG
//            Bot.SendTextMessageAsync(config.getGroupKey(GROUP_TYPE.GROUP_TYPE_CLAN), strPrint);  // 클랜방
//#endif
        }

        // init methods...
        public async void telegramAPIAsync()
        {
            //Bot 에 대한 정보를 가져온다.
            var me = await Bot.GetMeAsync();
        }

        public void setTelegramEvent()
        {
            Bot.OnMessage += Bot_OnMessage;     // 이벤트를 추가해줍니다. 

            Bot.StartReceiving();               // 이 함수가 실행이 되어야 사용자로부터 메세지를 받을 수 있습니다.
        }

        public bool loadUserInfo()
        {
            bool bResult = true;

            // Define request parameters.
            String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
            String range = "클랜원 목록!C8:V";
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = request.Execute();
            if (response != null)
            {
                IList<IList<Object>> values = response.Values;
                if (values != null && values.Count > 0)
                {
                    foreach (var row in values)
                    {
                        bool isReflesh = false;

                        // 아테나에 등록되지 않은 유저
                        if (row[18].ToString() == "")
                            continue;

                        long userKey = Convert.ToInt64(row[18].ToString());
                        var userData = userDirector.getUserInfo(userKey);
                        if (userData.UserKey != 0)
                        {
                            // 이미 등록한 유저. 갱신한다.
                            isReflesh = true;
                        }

                        CUser user = new CUser();
                        user = setUserInfo(row, Convert.ToInt64(row[18].ToString()));

                        // 운영진일 경우
                        if (config.isAdmin(user.UserKey))
                        {
                            // 유저 타입을 관리자로
                            user.UserType = USER_TYPE.USER_TYPE_ADMIN;
                        }

                        // 개발자일 경우
                        if (user.UserKey == 23842788)
                        {
                            // 유저 타입을 개발자로
                            user.UserType = USER_TYPE.USER_TYPE_DEVELOPER;
                        }

                        if (isReflesh == false)
                            userDirector.addUserInfo(userKey, user);
                        else
                            userDirector.refleshUserInfo(userKey, user);
                    }
                }
                else
                {
                    bResult = false;
                }
            }
            else
            {
                bResult = false;
            }

            return bResult;
        }

        public bool loadData()
        {
            bool bResult = true;

            string FolderName = @"Data/Chatting/";
            System.IO.DirectoryInfo chatDir = new System.IO.DirectoryInfo(FolderName);
            if (chatDir.Exists == false)
                return false;

            foreach (System.IO.FileInfo File in chatDir.GetFiles())
            {
                if (File.Name.Contains("ChattingCount_") == true)
                {
                    string fileName = File.Name.Replace("ChattingCount_", "");
                    fileName = fileName.Replace(".txt", "");

                    long userKey = Convert.ToInt64(fileName);

                    CUser userInfo = userDirector.getUserInfo(userKey);
                    if (userInfo.UserKey > 0)
                    {
                        string value = System.IO.File.ReadAllText(File.FullName, Encoding.UTF8);

                        userInfo.chattingCount = Convert.ToUInt64(value);
                    }
                }
            }

            FolderName = @"Data/Noti/";
            System.IO.DirectoryInfo notiDir = new System.IO.DirectoryInfo(FolderName);
            if (notiDir.Exists == false)
                return false;

            foreach (System.IO.FileInfo File in notiDir.GetFiles())
            {
                if (File.Name.Contains("Noti_") == true)
                {
                    string fileName = File.Name.Replace("Noti_", "");
                    fileName = fileName.Replace(".txt", "");

                    long userKey = Convert.ToInt64(fileName);

                    CUser userInfo = userDirector.getUserInfo(userKey);

                    string[] memoValue = System.IO.File.ReadAllLines(FolderName + File.Name);
                    if (memoValue.Length > 0)
                    {
                        foreach (var elem in memoValue)
                        {
                            string[] notiData = elem.Split('|');

                            CPrivateNoti noti = new CPrivateNoti();
                            noti.Hour = Convert.ToInt32(notiData[0]);
                            noti.Minute = Convert.ToInt32(notiData[1]);
                            noti.UserID = notiData[2];
                            noti.Notice = notiData[3];

                            userInfo.addPrivateNoti(noti);
                        }
                    }
                }
            }

            FolderName = @"Data/Memo/";
            System.IO.DirectoryInfo memoDir = new System.IO.DirectoryInfo(FolderName);
            if (memoDir.Exists == false)
                return false;

            foreach (System.IO.FileInfo File in memoDir.GetFiles())
            {
                if (File.Name.Contains("Memo_") == true)
                {
                    string fileName = File.Name.Replace("Memo_", "");
                    fileName = fileName.Replace(".txt", "");

                    long userKey = Convert.ToInt64(fileName);

                    CUser userInfo = userDirector.getUserInfo(userKey);                    
                    
                    string[] memoValue = System.IO.File.ReadAllLines(FolderName + File.Name);
                    if (memoValue.Length > 0)
                    {
                        foreach (var elem in memoValue)
                        {
                            CMemo memo = new CMemo();
                            memo.Memo = elem.ToString();
                            userInfo.addMemo(memo);
                        }
                    }
                }
            }

            return bResult;
        }

        public CUser setUserInfo(IList<object> row, long userKey)
        {
            CUser user = new CUser();

            if (row.Count == 0)
                return user;
            
            user.UserKey = userKey;
            user.Name = row[0].ToString();
            user.MainBattleTag = row[1].ToString();
            user.SubBattleTag = row[2].ToString().Trim().Split(',');
            user.Tier = row[3].ToString();

            if (row[4].ToString() == "플렉스")
                user.Position |= POSITION.POSITION_FLEX;
            if (row[4].ToString().ToUpper().Contains("딜"))
                user.Position |= POSITION.POSITION_DPS;
            if (row[4].ToString().ToUpper().Contains("탱"))
                user.Position |= POSITION.POSITION_TANK;
            if (row[4].ToString().ToUpper().Contains("힐"))
                user.Position |= POSITION.POSITION_SUPP;

            string[] most = new string[3];
            most[0] = row[5].ToString();
            most[1] = row[6].ToString();
            most[2] = row[7].ToString();
            user.MostPick = most;

            user.OtherPick = row[8].ToString();
            user.Team = row[9].ToString();
            user.Youtube = row[10].ToString();
            user.Twitch = row[11].ToString();
            user.Info = row[12].ToString();

            // 운영진일 경우
            if (config.isAdmin(user.UserKey))
            {
                // 유저 타입을 관리자로
                user.UserType = USER_TYPE.USER_TYPE_ADMIN;
            }

            if (user.UserKey == 23842788)
            {
                // 유저 타입을 개발자로
                user.UserType = USER_TYPE.USER_TYPE_DEVELOPER;
            }

            return user;
        }

        public Tuple<int, string> referenceScore(string battleTag)
        {
            int score = 0;
            string tier = "";

            string[] strBattleTag = battleTag.Split('#');
            string strUrl = "http://playoverwatch.com/ko-kr/career/pc/" + strBattleTag[0] + "-" + strBattleTag[1];

            try
            {
                WebClient wc = new WebClient();
                wc.Encoding = Encoding.UTF8;

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

                string html = wc.DownloadString(strUrl);
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(html);

                string strScore = doc.DocumentNode.SelectSingleNode("//div[@class='competitive-rank']").InnerText;
                score = Convert.ToInt32(strScore);

                if (score == 0)
                {
                    tier = "Unranked";
                }
                else if (score >= 0 && score < 1500)
                {
                    tier = "브론즈";
                }
                else if (score >= 1500 && score < 2000)
                {
                    tier = "실버";
                }
                else if (score >= 2000 && score < 2500)
                {
                    tier = "골드";
                }
                else if (score >= 2500 && score < 3000)
                {
                    tier = "플래티넘";
                }
                else if (score >= 3000 && score < 3500)
                {
                    tier = "다이아";
                }
                else if (score >= 3500 && score < 4000)
                {
                    tier = "마스터";
                }
                else if (score >= 4000 && score <= 5000)
                {
                    tier = "그랜드마스터";
                }
            }
            catch
            {
                // 아무 작업 안함
            }

            Tuple<int, string> retTuple = Tuple.Create(score, tier);
            return retTuple;
        }

        // 쓰레드풀의 작업쓰레드가 지정된 시간 간격으로
        // 아래 이벤트 핸들러 실행
        public void timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            string strPrint = "";
            
            if (systemInfo.getNowDayOfWeek() == DayOfWeek.Sunday)
            {
                if (isLupinOfWeek == false && DateTime.Now.Hour == 23)
                {
                    var allUser = userDirector.getAllUserInfo();
                    Dictionary<long, ulong> dicChattingCount = new Dictionary<long, ulong>();
                    ulong totalCount = 0;

                    foreach (var iter in allUser)
                    {
                        dicChattingCount.Add(iter.Value.UserKey, iter.Value.chattingCount);
                        totalCount += iter.Value.chattingCount;
                    }

                    int rank = 1;
                    int afterRank = 1;
                    ulong afterValue = 0;
                    bool isContinue = false;
                    strPrint += "[ 금주의 채팅 순위 Top 10 ]\n============================\n";
                    foreach (KeyValuePair<long, ulong> item in dicChattingCount.OrderByDescending(key => key.Value))
                    {
                        if (rank > 10)
                            break;

                        if (item.Value == 0)
                            break;

                        var user = userDirector.getUserInfo(item.Key);
                        double value = Convert.ToDouble(item.Value);
                        double count = Convert.ToDouble(totalCount);

                        if (afterValue == item.Value)
                        {
                            if (isContinue == false)
                                afterRank--;

                            isContinue = true;
                            strPrint += afterRank.ToString() + ". " + user.Name + " - " + item.Value + "건 (" + Math.Round(value / count * 100.0, 2) + "%)\n";
                            rank++;
                        }
                        else
                        {
                            isContinue = false;
                            strPrint += rank.ToString() + ". " + user.Name + " - " + item.Value + "건 (" + Math.Round(value / count * 100.0, 2) + "%)\n";
                            rank++;
                            afterRank = rank;
                        }

                        afterValue = item.Value;
                    }

                    userDirector.resetChattingCount();
                    strPrint += "\n대화량 카운트가 초기화 됐습니다.";
                    isLupinOfWeek = true;

#if DEBUG
                    Bot.SendTextMessageAsync(config.getGroupKey(GROUP_TYPE.GROUP_TYPE_TEST), strPrint);  // 운영진방
#else
                    Bot.SendTextMessageAsync(config.getGroupKey(GROUP_TYPE.GROUP_TYPE_CLAN), strPrint);  // 클랜방
#endif
                }
                else if (DateTime.Now.Hour != 23)
                {
                    isLupinOfWeek = false;
                }
            }

            if (DateTime.Now.Hour == 8)
            {
                if (isGoodMorning == false)
                {
                    isGoodMorning = true;
                    strPrint += "굿모닝~ 오늘도 즐거운 하루 되세요~ :)";

#if DEBUG
                    Bot.SendTextMessageAsync(config.getGroupKey(GROUP_TYPE.GROUP_TYPE_TEST), strPrint);  // 운영진방
#else
                    Bot.SendTextMessageAsync(config.getGroupKey(GROUP_TYPE.GROUP_TYPE_CLAN), strPrint);  // 클랜방
#endif
                }
            }
            else
            {
                isGoodMorning = false;
            }

            var allUserInfo = userDirector.getAllUserInfo();

            foreach (var elem in allUserInfo)
            {
                var privateNoti = elem.Value.getPrivateNoti();

                if (privateNoti.Count > 0)
                {
                    int index = 0;

                    foreach (var noti in privateNoti)
                    {
                        // 세팅한 시간과 동일하면
                        if (DateTime.Now.Hour == noti.Hour && DateTime.Now.Minute == noti.Minute)
                        {
                            string userNoti = noti.Notice.ToString();

                            strPrint += "[알림] " + userNoti + " / @" + noti.UserID;

                            if (userDirector.getPrivateNoti(elem.Value.UserKey).Count > 0)
                            {
                                userDirector.RemoveNoti(elem.Value.UserKey, index);
#if DEBUG
                                Bot.SendTextMessageAsync(config.getGroupKey(GROUP_TYPE.GROUP_TYPE_TEST), strPrint);  // 운영진방
#else
                                Bot.SendTextMessageAsync(config.getGroupKey(GROUP_TYPE.GROUP_TYPE_CLAN), strPrint);  // 클랜방
#endif
                                strPrint = "";
                                index++;
                                break;
                            }
                        }

                        index++;
                    }
                }
            }

            // Define request parameters.
            String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
            String range = "클랜 공지!C17:C23";
            String updateRange = "클랜 공지!H16";
            String calendarUpdateRange = "클랜 공지!O4";
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
            SpreadsheetsResource.ValuesResource.GetRequest updateRequest = service.Spreadsheets.Values.Get(spreadsheetId, updateRange);
            SpreadsheetsResource.ValuesResource.GetRequest calendarUpdateRequest = service.Spreadsheets.Values.Get(spreadsheetId, calendarUpdateRange);

            ValueRange response = request.Execute();
            ValueRange updateResponse = updateRequest.Execute();
            ValueRange calendarUpdateResponse = calendarUpdateRequest.Execute();

            if (response.Values != null && updateResponse.Values != null)
            {
                IList<IList<Object>> values = response.Values;
                IList<IList<Object>> updateValues = updateResponse.Values;

                // 공지
                if (updateValues != null && updateValues.ToString() != "")
                {
                    if (values != null && values.Count > 0)
                    {
                        strPrint += "#공지사항\n\n";

                        foreach (var row in values)
                        {
                            strPrint += "* " + row[0] + "\n\n";
                        }
                    }

                    Notice.SetNotice(strPrint);

                    // Define request parameters.
                    ValueRange valueRange = new ValueRange();
                    valueRange.MajorDimension = "COLUMNS"; //"ROWS";//COLUMNS 

                    var oblist = new List<object>() { "" };
                    valueRange.Values = new List<IList<object>> { oblist };

                    SpreadsheetsResource.ValuesResource.UpdateRequest releaseRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, updateRange);

                    releaseRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                    UpdateValuesResponse releaseResponse = releaseRequest.Execute();
                    if (releaseResponse == null)
                    {
                        strPrint = "[ERROR] 시트를 업데이트 할 수 없습니다.";
                    }

                    const string notice = @"Function/Notice.jpg";
                    var fileName = notice.Split(Path.DirectorySeparatorChar).Last();
                    var fileStream = new FileStream(notice, FileMode.Open, FileAccess.Read, FileShare.Read);
#if DEBUG
                    Bot.SendPhotoAsync(config.getGroupKey(GROUP_TYPE.GROUP_TYPE_TEST), fileStream, strPrint);  // 운영진방
#else
                    Bot.SendPhotoAsync(config.getGroupKey(GROUP_TYPE.GROUP_TYPE_CLAN), fileStream, strPrint);  // 클랜방
#endif
                }
            }
            
            if (calendarUpdateResponse.Values != null)
            {
                CCalendarDirector calendarDirector = new CCalendarDirector();

                String calendarRange = "클랜 공지!I8:P31";
                SpreadsheetsResource.ValuesResource.GetRequest calendarRequest = service.Spreadsheets.Values.Get(spreadsheetId, calendarRange);

                ValueRange calendarResponse = calendarRequest.Execute();
                if (calendarResponse != null)
                {
                    IList<IList<Object>> values = calendarResponse.Values;
                    if (values != null && values.Count > 0)
                    {
                        // 날짜
                        for (int index = 0; index < 24; index += 4)
                        {
                            for (int week = 0; week < 7; week++)
                            {
                                int calumn = 0;
                                CCalendar calendar = new CCalendar();
                                var row = values[index + calumn++];

                                string[] dateSplit = row[week].ToString().Split('/');

                                DateTime dateTime;
                                if (System.DateTime.Now.Year == 2019 && Convert.ToInt32(dateSplit[0]) == 1)
                                {
                                    dateTime = new DateTime(2020, Convert.ToInt32(dateSplit[0]), Convert.ToInt32(dateSplit[1]));
                                }
                                else
                                {
                                    dateTime = new DateTime(2019, Convert.ToInt32(dateSplit[0]), Convert.ToInt32(dateSplit[1]));
                                }

                                calendar.Time = dateTime;
                                calendar.Todo = new string[3];

                                row = values[index + calumn++];
                                if (row[week].ToString() == "")
                                    continue;

                                // 일정1
                                int iter = 0;
                                calendar.Todo[iter++] = row[week].ToString();

                                // 일정2
                                row = values[index + calumn++];
                                calendar.Todo[iter++] = row[week].ToString();

                                // 일정3
                                row = values[index + calumn];
                                calendar.Todo[iter] = row[week].ToString();

                                calendarDirector.addCalendar(calendar);
                            }
                        }
                    }
                }


                var calendarPrint = calendarDirector.getCalendar();
                int month = 0;
                bool isTitle = false;

                foreach (var elem in calendarPrint)
                {
                    if (month != elem.Value.Time.Month)
                    {
                        month = elem.Value.Time.Month;
                        isTitle = true;
                    }

                    DateTime nowTime = System.DateTime.Now;
                    DateTime todayTime = new DateTime(nowTime.Year, nowTime.Month, nowTime.Day);
                    DateTime nextTime = nowTime.AddDays(14);    // 기본값 2주

                    if (nextTime < elem.Value.Time)
                        continue;

                    if (todayTime > elem.Value.Time)
                        continue;

                    if (isTitle == true)
                    {
                        strPrint += "\n[ " + elem.Value.Time.Month + "월 일정 ]\n====================\n";
                        isTitle = false;
                    }

                    delegateDayOfWeek getDayOfWeek = (DayOfWeek week) =>
                    {
                        switch (week)
                        {
                            case DayOfWeek.Sunday:
                                return "일";
                            case DayOfWeek.Monday:
                                return "월";
                            case DayOfWeek.Tuesday:
                                return "화";
                            case DayOfWeek.Wednesday:
                                return "수";
                            case DayOfWeek.Thursday:
                                return "목";
                            case DayOfWeek.Friday:
                                return "금";
                            case DayOfWeek.Saturday:
                                return "토";
                        }

                        return "없음";
                    };

                    strPrint += elem.Value.Time.Day + "일(" + getDayOfWeek(elem.Value.Time.DayOfWeek) + ") : ";
                    bool isFirst = true;
                    foreach (var todo in elem.Value.Todo)
                    {
                        if (todo == "")
                            continue;

                        if (isFirst == false)
                            strPrint += " / ";

                        strPrint += todo;
                        isFirst = false;
                    }

                    strPrint += "\n";
                }

                // Define request parameters.
                ValueRange valueRange = new ValueRange();
                valueRange.MajorDimension = "COLUMNS"; //"ROWS";//COLUMNS 

                var oblist = new List<object>() { "" };
                valueRange.Values = new List<IList<object>> { oblist };

                SpreadsheetsResource.ValuesResource.UpdateRequest releaseRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, calendarUpdateRange);

                releaseRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                UpdateValuesResponse releaseResponse = releaseRequest.Execute();
                if (releaseResponse == null)
                {
                    strPrint = "[ERROR] 시트를 업데이트 할 수 없습니다.";
                }

                if (strPrint != "")
                {
                    const string notice = @"Function/Calendar.png";
                    var fileName = notice.Split(Path.DirectorySeparatorChar).Last();
                    var fileStream = new FileStream(notice, FileMode.Open, FileAccess.Read, FileShare.Read);
#if DEBUG
                    Bot.SendPhotoAsync(config.getGroupKey(GROUP_TYPE.GROUP_TYPE_TEST), fileStream, strPrint);  // 운영진방
#else
                    Bot.SendPhotoAsync(config.getGroupKey(GROUP_TYPE.GROUP_TYPE_CLAN), fileStream, strPrint);  // 클랜방
#endif
                }
            }
        }

        delegate string delegateDayOfWeek(DayOfWeek week);

        //Events...
        // Telegram...
        private async void Bot_OnMessage(object sender, Telegram.Bot.Args.MessageEventArgs e)
        {
            var varMessage = e.Message;
            if (varMessage == null)
                return;

            // 업로드일 경우
            if (varMessage.Caption == "/up")
                varMessage.Text = "/up";

            DateTime convertTime = varMessage.Date.AddHours(9);
            if (convertTime < systemInfo.GetStartTimeToDate())
            {
                return;
            }

            // 메시지 정보 추출
            string strFirstName = varMessage.From.FirstName;
            string strLastName = varMessage.From.LastName;
            string strUserID = varMessage.From.Username;
            int iMessageID = varMessage.MessageId;
            long senderKey = varMessage.From.Id;
            DateTime time = convertTime;

            // 차단된 유저는 이용할 수 없다.
            if (userDirector.isBlockUser(senderKey) == true)
                return;

            // 대화량 누적
            if (config.getGroupType(varMessage.Chat.Id) == GROUP_TYPE.GROUP_TYPE_CLAN)
            {
                userDirector.increaseChattingCount(senderKey);
            }

            // 입장 메시지 일 경우
            if (varMessage.Type == MessageType.ChatMembersAdded)
            {
                if (varMessage.Chat.Id == config.getGroupKey(GROUP_TYPE.GROUP_TYPE_GUIDE))  // 사전안내방
                {
                    varMessage.Text = "/안내";
                }
#if !DEBUG
                else if (varMessage.Chat.Id == config.getGroupKey(GROUP_TYPE.GROUP_TYPE_CLAN))      // 본방
#else
                else if (varMessage.Chat.Id == config.getGroupKey(GROUP_TYPE.GROUP_TYPE_TEST))   // 테스트방
#endif
                {
                    string strInfo = "";

                    strInfo += "\n안녕하세요.\n";
                    strInfo += "서로의 삶에 힘이 되는 오버워치 클랜,\n";
                    strInfo += "'클리앙 딜리셔스 팀'에 오신 것을 환영합니다.\n\n";
                    strInfo += "저는 팀의 운영 봇인 아테나입니다.\n";
                    strInfo += "\n";
                    strInfo += "클랜 생활에 불편하신 점이 있으시거나\n";
                    strInfo += "건의사항, 문의사항이 있으실 때는\n";
                    strInfo += "네이티브(@nativehyun)\n";
                    strInfo += "봄의캐롤(@rsmini),\n";
                    strInfo += "냉각콜라(@Seungman),\n";
                    strInfo += "Sugar Belle(@sugarbell) 에게\n";
                    strInfo += "문의해주세요.\n\n";
                    strInfo += "클랜원들의 편의를 위한\n";
                    strInfo += "저, 아테나의 기능을 확인하시려면\n";
                    strInfo += "/도움말 을 입력해주세요.\n";
                    strInfo += "아테나에 대해 문의사항이 있으실 때는\n";
                    strInfo += "휴린(@hyulin)에게 문의해주세요.\n";
                    strInfo += "\n";
                    strInfo += "우리 클랜의 모든 일정관리 및 운영은\n";
                    strInfo += "통합문서를 통해 확인 하실 수 있습니다.\n";
                    strInfo += "(https://goo.gl/nurbLT [딜리셔스.kr])\n";
                    strInfo += "\n";
                    strInfo += "저희 CDT에서 즐거운 오버워치 생활,\n";
                    strInfo += "그리고 더 나아가 즐거운 라이프를\n";
                    strInfo += "즐기셨으면 좋겠습니다.\n\n";
                    strInfo += "잘 부탁드립니다 :)\n";

                    const string record = @"Function/Logo.jpg";
                    var fileName = record.Split(Path.DirectorySeparatorChar).Last();
                    var fileStream = new FileStream(record, FileMode.Open, FileAccess.Read, FileShare.Read);
                    await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream, strInfo);

                    const string strCDTInfo06 = @"CDT_Info/CDT_Info_6.png";
                    var fileName06 = strCDTInfo06.Split(Path.DirectorySeparatorChar).Last();
                    var fileStream06 = new FileStream(strCDTInfo06, FileMode.Open, FileAccess.Read, FileShare.Read);
                    await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream06, "");

                    return;
                }
                else
                {
                    return;
                }
            }

            // 텍스트가 없으면 넘어간다.
            if (varMessage.Text == null)
                return;

            // 명령어, 서브명령어 분리
            string strMassage = varMessage.Text;
            string strUserName = varMessage.From.FirstName + varMessage.From.LastName;
            string strCommend = "";
            string strContents = "";
            //bool isCommand = false;

            // 명령어인지 아닌지 구분
            if (strMassage.Substring(0, 1) == "/")
            {
                //isCommand = true;

                // 명령어와 서브명령어 구분
                if (strMassage.IndexOf(" ") == -1)
                {
                    strCommend = strMassage;
                }
                else
                {
                    strCommend = strMassage.Substring(0, strMassage.IndexOf(" "));
                    strContents = strMassage.Substring(strMassage.IndexOf(" ") + 1, strMassage.Count() - strMassage.IndexOf(" ") - 1);
                }

                // 미등록 유저는 사용할 수 없다.
                if (strCommend != "/등록" && strCommend != "/안내" && userDirector.getUserInfo(senderKey).UserKey == 0)
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 아테나에 등록되지 않은 유저입니다.\n등록을 하시려면 /등록 [본계정배틀태그]를 입력해주세요.", ParseMode.Default, false, false, iMessageID);
                    CLog.WriteLog(varMessage.Chat.Id, senderKey, strUserName, "[ERROR] 아테나에 등록되지 않은 유저입니다.\n등록을 하시려면 /등록 [본계정배틀태그]를 입력해주세요.", strCommend, strContents);
                    return;
                }
            }

            // CDT 관련방 아니면 동작하지 않도록
            if (varMessage.Chat.Id != config.getGroupKey(GROUP_TYPE.GROUP_TYPE_CLAN) &&     // 본방
                varMessage.Chat.Id != config.getGroupKey(GROUP_TYPE.GROUP_TYPE_TEST) &&     // 테스트방
                config.isDeveloper(senderKey) == false)
            {
                if (varMessage.Chat.Id == config.getGroupKey(GROUP_TYPE.GROUP_TYPE_GUIDE))     // 사전안내방
                {
                    // 사전안내방에서는 등록, 스크림, 조사, 안내만 가능
                    if (strCommend != "/등록" && strCommend != "/스크림" && strCommend != "/조사" && strCommend != "/안내")
                    {
                        return;
                    }
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 사용할 수 없는 대화방입니다.", ParseMode.Default, false, false, iMessageID);
                    CLog.WriteLog(varMessage.Chat.Id, senderKey, "", "[ERROR] 사용할 수 없는 대화방입니다.", "", "");
                    return;
                }
            }

            if (userDirector.getUserInfo(senderKey).UserKey > 0)
            {
                // 본방에 입력된 메시지를 각 유저 정보에 입력
#if DEBUG
                userDirector.addMessage(senderKey, strMassage, time);
#else
                if (senderKey != 0 && strMassage != "" && varMessage.Chat.Id == config.getGroupKey(GROUP_TYPE.GROUP_TYPE_CLAN))
                    userDirector.addMessage(senderKey, strMassage, time);
#endif
            }
            else
            {
                if (strCommend != "/안내" && strCommend != "/등록")
                    return;
            }
            /*
            // 명령어가 아닐 경우와 사전안내방이 아닌 경우
            if (isCommand == false && varMessage.Chat.Id != config.getGroupKey(GROUP_TYPE.GROUP_TYPE_GUIDE))
            {
                bool isMainRoom = false;

                if (varMessage.Chat.Id == config.getGroupKey(GROUP_TYPE.GROUP_TYPE_CLAN))
                    isMainRoom = true;

                Tuple<string, string, bool> tuple = naturalLanguage.morphemeProcessor(strMassage, userDirector.getMessage(senderKey), isMainRoom);

                // 대화
                if (tuple.Item1 != "" && tuple.Item3 == false)
                {
                    if (tuple.Item1.ToString() != "")
                    {
                        CLog.WriteLog(varMessage.Chat.Id, senderKey, strUserName, strMassage, tuple.Item1.ToString(), "");
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, tuple.Item1.ToString(), ParseMode.Default, false, false, iMessageID);
                    }
                }
                else if (tuple.Item3 == true)
                {
                    if (tuple.Item1.ToString() != "")
                    {
                        strCommend = tuple.Item1;
                        strContents = tuple.Item2;

                        CLog.WriteLog(varMessage.Chat.Id, senderKey, strUserName, strMassage, strCommend, strContents);
                    }
                }
                else if (tuple.Item1 == "" && tuple.Item2 == "" && tuple.Item3 == false)
                {
                    if (varMessage.ReplyToMessage != null && varMessage.ReplyToMessage.From.Username != null)
                    {
#if DEBUG
                        if (varMessage.ReplyToMessage.From.Username.ToString().Contains("CDT_Noti_Test_Bot") == true)
#else
                        if (varMessage.ReplyToMessage.From.Username.ToString().Contains("CDT_Noti_Bot") == true)
#endif
                        {
                            CLog.WriteLog(varMessage.Chat.Id, senderKey, strUserName, strMassage, tuple.Item1.ToString(), "");
                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, naturalLanguage.replyCall(strMassage), ParseMode.Default, false, false, iMessageID);
                        }
                    }
                }
            }
            else
            {
                CLog.WriteLog(varMessage.Chat.Id, senderKey, strUserName, strMassage, strCommend, strContents);
            }
            */
            string strPrint = "";

            //========================================================================================
            // 도움말
            //========================================================================================
            if (strCommend == "/도움말" || strCommend == "/help" || strCommend == "/help@CDT_Noti_Bot")
            {
                strPrint += "==================================\n";
                strPrint += "[ 아테나 v2.3 ]\n[ Clien Delicious Team Notice Bot ]\n\n";
                strPrint += "/공지 : 클랜 공지사항을 출력합니다.\n";
                //strPrint += "/문의 [내용] : 문의사항을 등록합니다.\n";
                strPrint += "/일정 : 이번 달 클랜 일정을 확인합니다.\n";
                strPrint += "/등록 [본 계정 배틀태그] : 아테나에 등록 합니다.\n";
                strPrint += "/조회 [검색어] : 클랜원을 조회합니다.\n";
                strPrint += "               (검색범위 : 대화명, 배틀태그, 부계정)\n";
                strPrint += "/영상 [날짜] : 플레이 영상을 조회합니다. (/영상 20181006)\n";
                strPrint += "/검색 [검색어] : 포지션, 모스트별로 클랜원을 검색합니다.\n";
                strPrint += "/스크림 : 현재 모집 중인 스크림의 참가자를 출력합니다.\n";
                strPrint += "/스크림 [요일] : 현재 모집 중인 스크림에 참가신청합니다.\n";
                strPrint += "/스크림 취소 : 신청한 스크림에 참가를 취소합니다.\n";
                strPrint += "/디비전 : 현재 모집 중인 오픈디비전의 참가자를 출력합니다.\n";
                strPrint += "/디비전 [요일] : 현재 모집 중인 오픈디비전에 참가신청합니다.\n";
                strPrint += "/디비전 취소 : 신청한 오픈디비전에 참가를 취소합니다.\n";
                strPrint += "/리그 : 현재 모집 중인 리그의 참가자를 출력합니다.\n";
                strPrint += "/리그 [요일] : 현재 모집 중인 리그에 참가신청합니다.\n";
                strPrint += "/리그 취소 : 신청한 리그에 참가를 취소합니다.\n";
                strPrint += "/조사 : 현재 진행 중인 일정 조사를 출력합니다.\n";
                strPrint += "/조사 [요일] : 현재 진행 중인 일정 조사에 체크합니다.\n";
                strPrint += "/모임 : 모임 공지와 참가자를 출력합니다.\n";
                strPrint += "/모임 투표 : 모임 날짜와 장소 투료 현황을 확인합니다..\n";
                strPrint += "/모임 [날짜] [장소] : 날짜와 장소를 투표합니다.\n";
                strPrint += "               (/모임 1 1 , /모임 12 123 등)\n";
                strPrint += "/모임 참가 : 모임에 참가 신청합니다.\n";
                strPrint += "/모임 취소 : 모임에 참가 신청을 취소합니다.\n";
                strPrint += "/투표 : 현재 진행 중인 투표를 출력합니다.\n";
                strPrint += "/투표 [숫자] : 현재 진행 중인 투표에 투표합니다.\n";
                strPrint += "/투표 결과 : 현재 진행 중인 투표의 결과를 출력합니다.\n";
                strPrint += "/기록 : 클랜 명예의 전당을 조회합니다.\n";
                strPrint += "/기록 [숫자] : 명예의 전당 상세내용을 조회합니다.\n";
                strPrint += "/뽑기 [항목1] [항목2] [항목3] ... : 하나를 뽑습니다.\n";
                strPrint += "/날씨 [지역] : 현재 날씨를 출력합니다.\n";
                strPrint += "/알림 : 현재 저장된 개인 알림 리스트를 출력합니다.\n";
                strPrint += "/알림 [시간] [내용] : 설정한 시간에 알림을 설정합니다.\n";
                strPrint += "/알림 제거 [시간] : 알림을 제거합니다.\n";
                strPrint += "/메모 : 현재 저장된 개인 메모 리스트를 출력합니다.\n";
                strPrint += "/메모 [내용] : 메모를 저장합니다..\n";
                strPrint += "/메모 제거 [숫자] : 메모를 제거합니다.\n";
                strPrint += "/순위 : 대화량 순위를 출력합니다.\n";
                strPrint += "/안내 : 팀 안내 메시지를 출력합니다.\n";
                strPrint += "/상태 : 현재 봇 상태를 출력합니다. 대답이 없으면 이상.\n";
                strPrint += "==================================\n";

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
            }
            //========================================================================================
            // 등록
            //========================================================================================
            else if (strCommend == "/등록")
            {
                if (strContents == "")
                {
                    strPrint += "[SYSTEM] 사용자 등록을 하려면\n/등록 [본 계정 배틀태그] 로 등록해주세요.\n(ex: /등록 휴린#3602)";
                }
                else
                {
                    string battleTag = strContents;

                    // Define request parameters.
                    String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                    String range = "클랜원 목록!C8:V";
                    SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                    ValueRange response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            int index = 0;
                            int searchIndex = 0;
                            int searchCount = 0;
                            bool isReflesh = false;
                            CUser user = new CUser();

                            foreach (var row in values)
                            {
                                // 검색 성공
                                if (battleTag == row[1].ToString())
                                {
                                    long userKey = 0;
                                    searchCount++;
                                    searchIndex = index;

                                    if (row[18].ToString() != "")
                                    {
                                        // 이미 값이 있으므로 갱신한다.
                                        userKey = Convert.ToInt64(row[18].ToString());
                                        isReflesh = true;
                                    }
                                    else
                                    {
                                        userKey = senderKey;
                                    }

                                    user.UserKey = userKey;
                                    user.Name = row[0].ToString();
                                    user.MainBattleTag = row[1].ToString();
                                    user.SubBattleTag = row[2].ToString().Trim().Split(',');

                                    if (row[4].ToString() == "플렉스")
                                        user.Position |= POSITION.POSITION_FLEX;
                                    if (row[4].ToString().ToUpper().Contains("딜"))
                                        user.Position |= POSITION.POSITION_DPS;
                                    if (row[4].ToString().ToUpper().Contains("탱"))
                                        user.Position |= POSITION.POSITION_TANK;
                                    if (row[4].ToString().ToUpper().Contains("힐"))
                                        user.Position |= POSITION.POSITION_SUPP;

                                    string[] most = new string[3];
                                    most[0] = row[5].ToString();
                                    most[1] = row[6].ToString();
                                    most[2] = row[7].ToString();
                                    user.MostPick = most;
                                    user.OtherPick = row[8].ToString();

                                    user.Team = row[9].ToString();
                                    user.Youtube = row[10].ToString();
                                    user.Twitch = row[11].ToString();
                                    user.Info = row[12].ToString();

                                    user.Monitor = row[13].ToString();
                                    user.HeadSet = row[14].ToString();
                                    user.Keyboard = row[15].ToString();
                                    user.Mouse = row[16].ToString();
                                    user.eDPI = row[17].ToString();
                                }
                                else
                                {
                                    index++;
                                }
                            }

                            if (searchCount == 0)
                            {
                                strPrint += "[ERROR] 배틀태그를 검색할 수 없습니다.";
                                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                                return;
                            }
                            else if (searchCount > 1)
                            {
                                strPrint += "[ERROR] 검색 결과가 2개 이상입니다. 배틀태그를 확인해주세요.";
                                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                                return;
                            }
                            else if (searchCount < 0)
                            {
                                strPrint += "[ERROR] 알 수 없는 문제";
                                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                                return;
                            }
                            else if (isReflesh == true)
                            {
                                if (userDirector.refleshUserInfo(user.UserKey, user) == true)
                                {
                                    strPrint += "[SUCCESS] 갱신 완료됐습니다.";
                                }
                            }

                            range = "클랜원 목록!U" + (8 + searchIndex);

                            // Define request parameters.
                            ValueRange valueRange = new ValueRange();
                            valueRange.MajorDimension = "COLUMNS"; //"ROWS";//COLUMNS 

                            var oblist = new List<object>() { senderKey };
                            valueRange.Values = new List<IList<object>> { oblist };

                            SpreadsheetsResource.ValuesResource.UpdateRequest updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);

                            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                            UpdateValuesResponse updateResponse = updateRequest.Execute();

                            if (updateResponse == null)
                            {
                                strPrint += "[ERROR] 시트를 업데이트 할 수 없습니다.";
                            }
                            else
                            {
                                if (strPrint == "")
                                    strPrint += "[SUCCESS] 등록이 완료됐습니다.";

                                if (isReflesh == false)
                                    userDirector.addUserInfo(senderKey, user);
                            }
                        }
                    }
                }

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
            }
            //========================================================================================
            // 공지사항
            //========================================================================================
            else if (strCommend == "/공지")
            {
                // Define request parameters.
                String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                String range = "클랜 공지!C17:C23";
                SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                ValueRange response = request.Execute();
                if (response != null)
                {
                    IList<IList<Object>> values = response.Values;
                    if (values != null && values.Count > 0)
                    {
                        strPrint += "#공지사항\n\n";

                        foreach (var row in values)
                        {
                            strPrint += "* " + row[0] + "\n\n";
                        }
                    }
                }

                if (strPrint != "")
                {
                    const string notice = @"Function/Logo.jpg";
                    var fileName = notice.Split(Path.DirectorySeparatorChar).Last();
                    var fileStream = new FileStream(notice, FileMode.Open, FileAccess.Read, FileShare.Read);
                    await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream, strPrint, ParseMode.Default, false, iMessageID);
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 공지가 등록되지 않았습니다.", ParseMode.Default, false, false, iMessageID);
                }
            }
            //========================================================================================
            // 일정
            //========================================================================================
            else if (strCommend == "/일정")
            {
                CCalendarDirector calendarDirector = new CCalendarDirector();
                int nextDay = 0;

                int checkNum = 0;
                if (int.TryParse(strContents, out checkNum) == true)
                {
                    nextDay = Convert.ToInt32(strContents);
                }
                else if (strContents == "")
                {
                    nextDay = 14;
                }
                else if (strContents == "전체")
                {
                    nextDay = 100;
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 기간을 잘못 입력했습니다.", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                // Define request parameters.
                String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                String range = "클랜 공지!I8:P31";
                SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                ValueRange response = request.Execute();
                if (response != null)
                {
                    IList<IList<Object>> values = response.Values;
                    if (values != null && values.Count > 0)
                    {
                        // 날짜
                        for (int index = 0; index < 24; index += 4)
                        {
                            for (int week = 0; week < 7; week++)
                            {
                                int calumn = 0;
                                CCalendar calendar = new CCalendar();
                                var row = values[index + calumn++];

                                string[] dateSplit = row[week].ToString().Split('/');

                                DateTime dateTime;
                                if (System.DateTime.Now.Year == 2019 && Convert.ToInt32(dateSplit[0]) == 1)
                                {
                                    dateTime = new DateTime(2020, Convert.ToInt32(dateSplit[0]), Convert.ToInt32(dateSplit[1]));
                                }
                                else
                                {
                                    dateTime = new DateTime(2019, Convert.ToInt32(dateSplit[0]), Convert.ToInt32(dateSplit[1]));
                                }

                                calendar.Time = dateTime;
                                calendar.Todo = new string[3];

                                row = values[index + calumn++];
                                if (row[week].ToString() == "")
                                    continue;

                                // 일정1
                                int iter = 0;
                                calendar.Todo[iter++] = row[week].ToString();

                                // 일정2
                                row = values[index + calumn++];
                                calendar.Todo[iter++] = row[week].ToString();

                                // 일정3
                                row = values[index + calumn];
                                calendar.Todo[iter] = row[week].ToString();

                                calendarDirector.addCalendar(calendar);
                            }
                        }
                    }
                }


                var calendarPrint = calendarDirector.getCalendar();
                int month = 0;
                bool isTitle = false;

                if (nextDay == 0)
                    strPrint += "* 오늘의 일정 *\n";
                else if (nextDay == 100)
                    strPrint += "* 현재 등록된 전체 일정 *\n";
                else if (nextDay != 100)
                    strPrint += "* 현재 날짜부터 " + nextDay + "일 간의 일정 *\n";

                foreach (var elem in calendarPrint)
                {
                    if (month != elem.Value.Time.Month)
                    {
                        month = elem.Value.Time.Month;
                        isTitle = true;
                    }

                    DateTime nowTime = System.DateTime.Now;
                    DateTime todayTime = new DateTime(nowTime.Year, nowTime.Month, nowTime.Day);
                    DateTime nextTime = nowTime.AddDays(nextDay);

                    if (nextTime < elem.Value.Time)
                        continue;

                    if (todayTime > elem.Value.Time)
                        continue;

                    if (isTitle == true)
                    {
                        strPrint += "\n[ " + elem.Value.Time.Month + "월 일정 ]\n====================\n";
                        isTitle = false;
                    }

                    delegateDayOfWeek getDayOfWeek = (DayOfWeek week) =>
                    {
                        switch (week)
                        {
                            case DayOfWeek.Sunday:
                                return "일";
                            case DayOfWeek.Monday:
                                return "월";
                            case DayOfWeek.Tuesday:
                                return "화";
                            case DayOfWeek.Wednesday:
                                return "수";
                            case DayOfWeek.Thursday:
                                return "목";
                            case DayOfWeek.Friday:
                                return "금";
                            case DayOfWeek.Saturday:
                                return "토";
                        }

                        return "없음";
                    };

                    strPrint += elem.Value.Time.Day + "일(" + getDayOfWeek(elem.Value.Time.DayOfWeek) + ") : ";

                    bool isFirst = true;
                    foreach (var todo in elem.Value.Todo)
                    {
                        if (todo == "")
                            continue;

                        if (isFirst == false)
                            strPrint += " / ";

                        strPrint += todo;
                        isFirst = false;
                    }

                    strPrint += "\n";
                }

                if (strPrint != "")
                {
                    const string notice = @"Function/Calendar.png";
                    var fileName = notice.Split(Path.DirectorySeparatorChar).Last();
                    var fileStream = new FileStream(notice, FileMode.Open, FileAccess.Read, FileShare.Read);
                    await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream, strPrint, ParseMode.Default, false, iMessageID);
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 일정이 등록되지 않았습니다.", ParseMode.Default, false, false, iMessageID);
                }
            }
            //========================================================================================
            // 조회
            //========================================================================================
            else if (strCommend == "/조회")
            {
                if (strContents == "")
                {
                    strPrint += "[ERROR] 조회 대상이 없습니다.";
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                }
                else
                {
                    // Define request parameters.
                    String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                    String range = "클랜원 목록!C8:V";
                    SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                    ValueRange response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            bool bContinue = false;
                            string[] strList = new string[5];
                            int iIndex = 0;

                            foreach (var row in values)
                            {
                                if (row[0].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[1].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[2].ToString().ToUpper().Contains(strContents.ToUpper()))
                                {
                                    if (iIndex++ > 2)
                                    {
                                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 검색 결과가 너무 많습니다. (3건 초과)\n검색어를 다시 입력해주세요.", ParseMode.Default, false, false, iMessageID);
                                        return;
                                    }
                                }
                            }

                            foreach (var row in values)
                            {
                                bool isSubAccount = false;
                                bool isSearch = false;
                                string strUrl = "";
                                string battleTag = "";
                                string mainBattleTag = "";

                                if (row[0].ToString().ToUpper().Contains(strContents.ToUpper()))
                                {
                                    isSubAccount = false;
                                    isSearch = true;
                                }

                                if (row[1].ToString().ToUpper().Contains(strContents.ToUpper()))
                                {
                                    isSubAccount = false;
                                    isSearch = true;
                                }

                                if (row[2].ToString().ToUpper().Contains(strContents.ToUpper()))
                                {
                                    isSubAccount = true;
                                    isSearch = true;
                                }

                                if (isSearch == true && isSubAccount == false)
                                {
                                    string[] strBattleTag = row[1].ToString().Split('#');
                                    battleTag = strBattleTag[0] + "#" + strBattleTag[1];
                                    mainBattleTag = row[1].ToString();
                                    strUrl = "http://playoverwatch.com/ko-kr/career/pc/" + strBattleTag[0] + "-" + strBattleTag[1];
                                }
                                else if (isSearch == true && isSubAccount == true)
                                {
                                    string[] strSubAccount = row[2].ToString().Split(',');
                                    mainBattleTag = row[1].ToString();

                                    foreach (var acc in strSubAccount)
                                    {
                                        if (acc.ToString().ToUpper().Contains(strContents.ToUpper()))
                                        {
                                            string[] strBattleTag = acc.ToString().Split('#');
                                            battleTag = strBattleTag[0] + "#" + strBattleTag[1];
                                            strUrl = "http://playoverwatch.com/ko-kr/career/pc/" + strBattleTag[0] + "-" + strBattleTag[1];
                                            break;
                                        }
                                    }
                                }
                                else
                                {
                                    continue;
                                }

                                string strScore = "전적을 조회할 수 없습니다.";
                                string strTier = "전적을 조회할 수 없습니다.";

                                await Bot.SendTextMessageAsync(varMessage.Chat.Id, "'" + battleTag + "'의 전적을 조회 중입니다.\n잠시만 기다려주세요.", ParseMode.Default, false, false, iMessageID);

                                try
                                {
                                    WebClient wc = new WebClient();
                                    wc.Encoding = Encoding.UTF8;

                                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                                    ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

                                    string html = wc.DownloadString(strUrl);
                                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                                    doc.LoadHtml(html);

                                    strScore = doc.DocumentNode.SelectSingleNode("//div[@class='competitive-rank']").InnerText;

                                    if (Int32.Parse(strScore) >= 0 && Int32.Parse(strScore) < 1500)
                                    {
                                        strTier = "브론즈";
                                    }
                                    else if (Int32.Parse(strScore) >= 1500 && Int32.Parse(strScore) < 2000)
                                    {
                                        strTier = "실버";
                                    }
                                    else if (Int32.Parse(strScore) >= 2000 && Int32.Parse(strScore) < 2500)
                                    {
                                        strTier = "골드";
                                    }
                                    else if (Int32.Parse(strScore) >= 2500 && Int32.Parse(strScore) < 3000)
                                    {
                                        strTier = "플래티넘";
                                    }
                                    else if (Int32.Parse(strScore) >= 3000 && Int32.Parse(strScore) < 3500)
                                    {
                                        strTier = "다이아";
                                    }
                                    else if (Int32.Parse(strScore) >= 3500 && Int32.Parse(strScore) < 4000)
                                    {
                                        strTier = "마스터";
                                    }
                                    else if (Int32.Parse(strScore) >= 4000 && Int32.Parse(strScore) <= 5000)
                                    {
                                        strTier = "그랜드마스터";
                                    }
                                }
                                catch
                                {
                                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "'" + battleTag + "'의 전적을 조회할 수 없습니다.", ParseMode.Default, false, false, iMessageID);
                                }

                                if (bContinue == true)
                                {
                                    strPrint += "==================================\n";
                                }

                                strPrint += "* 티어 및 점수는 전적을 조회합니다. *\n\n";
                                strPrint += "[ " + row[0].ToString() + " ]\n";
                                strPrint += "- 조회 배틀태그 : " + battleTag + "\n";

                                if (row[9].ToString() != "")
                                    strPrint += "- 소속팀 : " + row[9].ToString() + "\n";
                                else
                                    strPrint += "- 소속팀 : 없음\n";

                                strPrint += "- 티어 : " + strTier + "\n";
                                strPrint += "- 점수 : " + strScore + "\n";
                                strPrint += "- 본 계정 배틀태그 : " + mainBattleTag + "\n";

                                if (row[2].ToString() != "")
                                    strPrint += "- 부 계정 배틀태그 : " + row[2].ToString() + "\n";

                                strPrint += "- 포지션 : " + row[4].ToString() + "\n";

                                strPrint += "- 모스트 : ";
                                if (row[5].ToString() != "")
                                    strPrint += row[5].ToString();
                                if (row[6].ToString() != "")
                                    strPrint += " / " + row[6].ToString();
                                if (row[7].ToString() != "")
                                    strPrint += " / " + row[7].ToString();
                                strPrint += "\n";

                                if (row[8].ToString() != "")
                                    strPrint += "- 이외 가능 픽 : " + row[8].ToString() + "\n";

                                if (row[12].ToString() != "")
                                    strPrint += "- 자기소개 : " + row[12].ToString() + "\n";

                                strPrint += "----------------------------------\n";

                                if (row[10].ToString() != "")
                                    strPrint += "- Youtube : " + row[10].ToString() + "\n";
                                if (row[11].ToString() != "")
                                    strPrint += "- Twitch : " + row[11].ToString() + "\n";

                                if (row[10].ToString() != "" || row[11].ToString() != "")
                                    strPrint += "----------------------------------\n";

                                bContinue = true;   // 한 명만 출력된다면 이 부분은 무시됨.
                            }
                        }
                    }

                    if (strPrint != "")
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                    }
                    else
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 클랜원을 찾을 수 없습니다.", ParseMode.Default, false, false, iMessageID);
                    }
                }
            }
            //========================================================================================
            // 영상
            //========================================================================================
            else if (strCommend == "/영상")
            {
                if (strContents == "")
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 날짜를 잘못 입력했습니다.\n(ex: /영상 201812 , /영상 20181215)", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                // Define request parameters.
                String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);

                if (strContents.Length == 4 || strContents.Length == 6)
                {
                    string year = "";
                    string month = "";
                    bool[] isPass = new bool[3];

                    if (strContents.Length >= 4)
                    {
                        year = strContents.Substring(0, 4);
                        isPass[0] = false;
                    }
                    else
                    {
                        isPass[0] = true;
                    }

                    if (strContents.Length >= 6)
                    {
                        month = strContents.Substring(4, 2);
                        isPass[1] = false;
                    }
                    else
                    {
                        isPass[1] = true;
                    }

                    if (year != "2018" && year != "2019")
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 날짜를 잘못 입력했습니다.\n(ex: /영상 201812 , /영상 20181215)", ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    String range = "영상 URL (" + year + ")!B5:G";
                    SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                    ValueRange response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            foreach (var row in values)
                            {
                                if (row.Count() != 6 || row[0].ToString() == "")
                                    continue;

                                string[] splitDate = row[0].ToString().Split('.');

                                if (isPass[0] == false && splitDate[0] == year)
                                {
                                    if (isPass[1] == false && splitDate[1] == month)
                                        strPrint += "[" + row[0].ToString() + "] " + row[1].ToString() + "\n";

                                    if (isPass[1] == true)
                                        strPrint += "[" + row[0].ToString() + "] " + row[1].ToString() + "\n";
                                }
                            }
                        }
                    }

                    strPrint += "\n/영상 날짜로 영상 주소를 조회하실 수 있습니다.\n";
                    strPrint += "(ex: /영상 20181006)";
                }
                else if (strContents.Length == 8)
                {
                    string year = strContents.Substring(0, 4);
                    string month = strContents.Substring(4, 2);
                    string day = strContents.Substring(6, 2);
                    bool bContinue = false;
                    string user = "";
                    string date = year + "." + month + "." + day;

                    String range = "영상 URL (" + year + ")!B5:G";
                    SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                    ValueRange response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            foreach (var row in values)
                            {
                                if (row.Count >= 6)
                                {
                                    if (row[0].ToString() == date)
                                    {
                                        bContinue = true;
                                    }
                                    else if (row[0].ToString() != "" && bContinue == true)
                                    {
                                        bContinue = false;
                                    }

                                    if (bContinue == true)
                                    {
                                        if (row[1].ToString() != "")
                                        {
                                            strPrint += "[ " + row[1].ToString() + " ]" + "\n";
                                        }

                                        if (row[3].ToString() == "")
                                        {
                                            if (row[4].ToString() != "")
                                            {
                                                strPrint += user + " (" + row[4].ToString() + ")" + " : " + row[5].ToString() + "\n";
                                            }
                                            else
                                            {
                                                strPrint += user + " : " + row[5].ToString() + "\n";
                                            }

                                        }
                                        else
                                        {
                                            if (row[4].ToString() != "")
                                            {
                                                strPrint += row[3].ToString() + " (" + row[4].ToString() + ")" + " : " + row[5].ToString() + "\n";
                                            }
                                            else
                                            {
                                                strPrint += row[3].ToString() + " : " + row[5].ToString() + "\n";
                                            }

                                            user = row[3].ToString();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 날짜를 잘못 입력했습니다.\n(ex: /영상 201812 , /영상 20181215)", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                if (strPrint != "")
                {
                    const string video = @"Function/Video.jpg";
                    var fileName = video.Split(Path.DirectorySeparatorChar).Last();
                    var fileStream = new FileStream(video, FileMode.Open, FileAccess.Read, FileShare.Read);

                    if (strPrint.Length > 1000)
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream, strPrint, ParseMode.Default, false, iMessageID);
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 영상을 찾을 수 없습니다.", ParseMode.Default, false, false, iMessageID);
                }
            }
            //========================================================================================
            // 검색
            //========================================================================================
            else if (strCommend == "/검색")
            {
                if (strContents == "")
                {
                    strPrint += "[ERROR] 검색 조건이 없습니다.";
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                }
                else
                {
                    // Define request parameters.
                    String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                    String range = "클랜원 목록!C8:V";
                    SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                    bool bResult = false;

                    ValueRange response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            strPrint += "[ '" + strContents + "' 검색 결과 ]\n";

                            foreach (var row in values)
                            {
                                if (strContents == "힐" || strContents == "딜" || strContents == "탱" || strContents == "플렉스")
                                {
                                    if (row[4].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[5].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[6].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[7].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[4].ToString() == "플렉스")
                                    {
                                        strPrint += row[0] + "(" + row[1] + ") : " + row[4].ToString() + " (";

                                        if (row[5].ToString() != "")
                                            strPrint += row[5].ToString();
                                        if (row[6].ToString() != "")
                                            strPrint += " / " + row[6].ToString();
                                        if (row[7].ToString() != "")
                                            strPrint += " / " + row[7].ToString();

                                        strPrint += ")\n";

                                        bResult = true;
                                    }
                                }
                                else
                                {
                                    if (row[4].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[5].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[6].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[7].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[9].ToString().ToUpper().Contains(strContents.ToUpper()))
                                    {
                                        strPrint += row[0] + "(" + row[1] + ") : " + row[4].ToString() + " (";

                                        if (row[9].ToString() != "")
                                            strPrint += "소속:" + row[9].ToString() + ") (";
                                        if (row[5].ToString() != "")
                                            strPrint += row[5].ToString();
                                        if (row[6].ToString() != "")
                                            strPrint += " / " + row[6].ToString();
                                        if (row[7].ToString() != "")
                                            strPrint += " / " + row[7].ToString();

                                        strPrint += ")\n";

                                        bResult = true;
                                    }
                                }
                            }
                        }
                    }

                    if (bResult == true)
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                    }
                    else
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 검색 결과가 없습니다.", ParseMode.Default, false, false, iMessageID);
                    }
                }
            }
            //========================================================================================
            // 모임
            //========================================================================================
            else if (strCommend == "/모임")
            {
                string title = "";

                // 타이틀
                String title_spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                String title_range = "모임!B2";
                SpreadsheetsResource.ValuesResource.GetRequest title_request = service.Spreadsheets.Values.Get(title_spreadsheetId, title_range);

                ValueRange title_response = title_request.Execute();
                if (title_response != null)
                {
                    IList<IList<Object>> values = title_response.Values;
                    if (values == null)
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 현재 예정된 모임이 없습니다.", ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    title = "[ " + values[0][0] + " ]\n\n";
                }

                // 모임 공지 출력
                if (strContents == "")
                {
                    // Define request parameters.
                    String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                    String range = "모임!P7:Q28";
                    SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                    ValueRange response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            strPrint += title;
                            strPrint += values[0][0] + " : " + values[0][1] + "\n";
                            strPrint += values[1][0] + " : " + values[1][1] + "\n";
                            strPrint += values[2][0] + " : " + values[2][1] + "\n\n";

                            if (values[4][0] != null && values[5][0] != null)
                            {
                                strPrint += "[ " + values[4][0] + " ]\n";
                                strPrint += values[5][0] + "\n\n";
                            }

                            if (values[12][0] != null && values[13][0] != null)
                            {
                                strPrint += "[ " + values[12][0] + " ]\n";
                                strPrint += values[13][0] + "\n\n";
                            }

                            if (values[20][0] != null && values[21][0] != null)
                            {
                                strPrint += "[ " + values[20][0] + " ]\n";
                                strPrint += values[21][0] + "\n\n";
                            }
                        }
                    }

                    // Define request parameters.
                    String user_spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                    String user_range = "모임!C6:O";
                    SpreadsheetsResource.ValuesResource.GetRequest user_request = service.Spreadsheets.Values.Get(user_spreadsheetId, user_range);

                    ValueRange user_response = user_request.Execute();
                    if (user_response != null)
                    {
                        IList<IList<Object>> values = user_response.Values;
                        if (values != null && values.Count > 0)
                        {
                            strPrint += "----------------------------------------\n";
                            strPrint += "* 참석 확정 : ";

                            int joinCount = 0;
                            bool isFirst = true;
                            foreach (var row in values)
                            {
                                if (row[0] != null && row[0].ToString() != "")
                                {
                                    if (row[9].ToString() == "O")
                                    {
                                        if (isFirst != true)
                                            strPrint += " , ";
                                        else
                                            isFirst = false;

                                        strPrint += row[0].ToString();
                                        joinCount++;
                                    }
                                }
                            }

                            strPrint += "\n----------------------------------------\n";
                            strPrint += "* 현재 " + joinCount + "명\n";
                            strPrint += "----------------------------------------\n";
                        }
                    }

                    if (strPrint != "")
                    {
                        const string meeting = @"Function/Meeting.jpg";
                        var fileName = meeting.Split(Path.DirectorySeparatorChar).Last();
                        var fileStream = new FileStream(meeting, FileMode.Open, FileAccess.Read, FileShare.Read);
                        await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream, strPrint, ParseMode.Default, false, iMessageID);

                        return;
                    }
                }
                // 모임 투표 출력
                else if (strContents == "투표")
                {
                    // Define request parameters.
                    String vote_spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                    String vote_range = "모임!D5:K5";
                    SpreadsheetsResource.ValuesResource.GetRequest vote_request = service.Spreadsheets.Values.Get(vote_spreadsheetId, vote_range);

                    ValueRange vote_response = vote_request.Execute();
                    if (vote_response != null)
                    {
                        IList<IList<Object>> values = vote_response.Values;
                        if (values != null && values.Count > 0)
                        {
                            strPrint += title;

                            var row = values[0];
                            if (row.Count <= 4)
                            {
                                await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 항목에 문제가 있습니다.", ParseMode.Default, false, false, iMessageID);
                                return;
                            }

                            int index = 1;
                            strPrint += "* 날짜\n";

                            // 날짜
                            for (int i = 0; i < 4; i++)
                            {
                                if (row[i] != null && row[i].ToString() != "")
                                {
                                    strPrint += "(" + index++ + ") " + row[i] + "\n";
                                }
                            }

                            index = 1;
                            strPrint += "\n* 장소\n";

                            // 장소
                            for (int i = 4; i < row.Count; i++)
                            {
                                strPrint += "(" + index++ + ") " + row[i] + "\n";
                            }

                            strPrint += "\n- 투표 방법\n/모임 1 3 , /모임 12 3 , /모임 12 123 등";
                        }
                    }

                    if (strPrint != "")
                    {
                        const string meeting = @"Function/Meeting.jpg";
                        var fileName = meeting.Split(Path.DirectorySeparatorChar).Last();
                        var fileStream = new FileStream(meeting, FileMode.Open, FileAccess.Read, FileShare.Read);
                        await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream, strPrint, ParseMode.Default, false, iMessageID);

                        return;
                    }
                }
                // 모임 참가 신청
                else if (strContents == "참가" || strContents == "취소")
                {
                    int index = 0;
                    int blankIndex = 0;
                    string[] vote = new string[8];
                    bool isVote = false;

                    // Define request parameters.
                    String vote_spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                    String vote_range = "모임!C6:L";
                    SpreadsheetsResource.ValuesResource.GetRequest vote_request = service.Spreadsheets.Values.Get(vote_spreadsheetId, vote_range);

                    ValueRange vote_response = vote_request.Execute();
                    if (vote_response != null)
                    {
                        IList<IList<Object>> values = vote_response.Values;
                        if (values != null && values.Count > 0)
                        {
                            foreach (var row in values)
                            {
                                if (row[0].ToString() == "")
                                {
                                    if (blankIndex == 0)
                                    {
                                        for (int i = 0; i < 8; i++)
                                        {
                                            vote[i] = row[i + 1].ToString();
                                        }
                                        
                                        blankIndex = index;
                                    }
                                }
                                else
                                {
                                    if (row[0].ToString() == strUserName)
                                    {
                                        isVote = true;
                                        blankIndex = index;
                                    }
                                }

                                index++;
                            }

                            if (isVote == false && blankIndex == 0)
                                blankIndex = index;
                        }
                    }

                    String join_spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                    string join_range = "모임!C" + (6 + blankIndex) + ":L";
                    ValueRange join_valueRange = new ValueRange();
                    join_valueRange.MajorDimension = "ROWS"; //"ROWS";//COLUMNS 

                    string join = "";

                    if (strContents == "참가")
                        join = "O";
                    else
                        join = "";

                    var oblist = new List<object>()
                    {
                        strUserName,
                        vote[0],
                        vote[1],
                        vote[2],
                        vote[3],
                        vote[4],
                        vote[5],
                        vote[6],
                        vote[7],
                        join
                    };

                    join_valueRange.Values = new List<IList<object>> { oblist };

                    SpreadsheetsResource.ValuesResource.UpdateRequest join_updateRequest = service.Spreadsheets.Values.Update(join_valueRange, join_spreadsheetId, join_range);

                    join_updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                    UpdateValuesResponse join_updateResponse = join_updateRequest.Execute();
                    if (join_updateResponse == null)
                    {
                        strPrint += "[ERROR] 시트를 업데이트 할 수 없습니다.";
                    }
                    else
                    {
                        strPrint += "[SYSTEM] 모임 조사를 완료했습니다.";
                    }
                }
                // 보기 투표
                else
                {
                    string[] vote = strContents.Split(' ');
                    if (vote.Count() <= 1 || vote.Count() >= 3)
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 보기 선택이 잘못 됐습니다.", ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    int checkNum = 0;
                    bool isNumber = int.TryParse(vote[0], out checkNum);
                    if (isNumber == false)
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 보기 선택이 잘못 됐습니다.", ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    int dateVote = Convert.ToInt32(vote[0]);
                    int placeVote = Convert.ToInt32(vote[1]);

                    if (dateVote > 1234 || placeVote > 1234)
                    { 
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 보기 선택이 잘못 됐습니다.", ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    string[] date = new string[4];
                    string[] place = new string[4];

                    date[0] = (dateVote / 1000).ToString();
                    dateVote -= (Convert.ToInt32(date[0]) * 1000);
                    date[1] = (dateVote / 100).ToString();
                    dateVote -= (Convert.ToInt32(date[1]) * 100);
                    date[2] = (dateVote / 10).ToString();
                    dateVote -= (Convert.ToInt32(date[2]) * 10);
                    date[3] = dateVote.ToString();

                    place[0] = (placeVote / 1000).ToString();
                    placeVote -= (Convert.ToInt32(place[0]) * 1000);
                    place[1] = (placeVote / 100).ToString();
                    placeVote -= (Convert.ToInt32(place[1]) * 100);
                    place[2] = (placeVote / 10).ToString();
                    placeVote -= (Convert.ToInt32(place[2]) * 10);
                    place[3] = placeVote.ToString();

                    string[] dateInput = { "", "", "", "" }; //new string[4];
                    string[] placeInput = { "", "", "", "" }; //new string[4];

                    for (int i = 0; i < 4; i++)
                    {
                        switch (date[i])
                        {
                            case "0":
                                break;
                            case "1":
                                dateInput[0] = "O";
                                break;
                            case "2":
                                dateInput[1] = "O";
                                break;
                            case "3":
                                dateInput[2] = "O";
                                break;
                            case "4":
                                dateInput[3] = "O";
                                break;
                            default:
                                await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 보기 선택이 잘못 됐습니다.", ParseMode.Default, false, false, iMessageID);
                                return;
                        }

                        switch (place[i])
                        {
                            case "0":
                                break;
                            case "1":
                                placeInput[0] = "O";
                                break;
                            case "2":
                                placeInput[1] = "O";
                                break;
                            case "3":
                                placeInput[2] = "O";
                                break;
                            case "4":
                                placeInput[3] = "O";
                                break;
                            default:
                                await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 보기 선택이 잘못 됐습니다.", ParseMode.Default, false, false, iMessageID);
                                return;
                        }
                    }

                    int index = 0;
                    int blankIndex = 0;
                    bool isVote = false;

                    // Define request parameters.
                    String scan_spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                    String scan_range = "모임!C6:L";
                    SpreadsheetsResource.ValuesResource.GetRequest scan_request = service.Spreadsheets.Values.Get(scan_spreadsheetId, scan_range);

                    ValueRange scan_response = scan_request.Execute();
                    if (scan_response != null)
                    {
                        IList<IList<Object>> values = scan_response.Values;
                        if (values != null && values.Count > 0)
                        {
                            foreach (var row in values)
                            {
                                if (row[0].ToString() == "")
                                {
                                    if (blankIndex == 0)
                                    {
                                        for (int i = 0; i < 8; i++)
                                        {
                                            vote[i] = row[i + 1].ToString();
                                        }

                                        blankIndex = index;
                                    }
                                }
                                else
                                {
                                    if (row[0].ToString() == strUserName)
                                    {
                                        isVote = true;
                                        blankIndex = index;
                                    }
                                }

                                index++;
                            }

                            if (isVote == false && blankIndex == 0)
                                blankIndex = index;
                        }
                    }

                    String vote_spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                    string vote_range = "모임!C" + (6 + blankIndex) + ":K";
                    ValueRange vote_valueRange = new ValueRange();
                    vote_valueRange.MajorDimension = "ROWS"; //"ROWS";//COLUMNS 

                    var oblist = new List<object>()
                    {
                        strUserName,
                        dateInput[0],
                        dateInput[1],
                        dateInput[2],
                        dateInput[3],
                        placeInput[0],
                        placeInput[1],
                        placeInput[2],
                        placeInput[3]
                    };

                    vote_valueRange.Values = new List<IList<object>> { oblist };

                    SpreadsheetsResource.ValuesResource.UpdateRequest vote_updateRequest = service.Spreadsheets.Values.Update(vote_valueRange, vote_spreadsheetId, vote_range);

                    vote_updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                    UpdateValuesResponse vote_updateResponse = vote_updateRequest.Execute();
                    if (vote_updateResponse == null)
                    {
                        strPrint += "[ERROR] 시트를 업데이트 할 수 없습니다.";
                    }
                    else
                    {
                        strPrint += "[SYSTEM] 모임 조사를 완료했습니다.";
                    }
                }

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
            }
            //========================================================================================
            // 투표
            //========================================================================================
            else if (strCommend == "/투표")
            {
                bool isAnonymous = false;

                // Define request parameters.
                String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                String range = "투표!B4:J";
                SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                ValueRange response = request.Execute();
                if (response != null)
                {
                    IList<IList<Object>> values = response.Values;
                    if (values != null && values.Count > 0)
                    {
                        CVoteDirector voteDirector = new CVoteDirector();

                        // 익명 여부
                        var value = values[0];
                        if (value[3].ToString() != "")
                        {
                            isAnonymous = true;
                        }

                        // 투표 내용
                        value = values[1];
                        string voteContents = value[0].ToString();
                        voteDirector.setVoteContents(voteContents);

                        if (voteContents != "")
                        {
                            // 투표 항목
                            value = values[11];
                            int index = 0;
                            int itemCount = 0;
                            foreach (var row in value)
                            {
                                string item = row.ToString();
                                if (item != "")
                                {
                                    CVoteItem voteItem = new CVoteItem();
                                    voteItem.AddItem(item);

                                    voteDirector.AddItem(voteItem);
                                    itemCount++;
                                }
                            }

                            // 투표자
                            index = 14;
                            while (true)
                            {
                                value = values[index++];

                                for (int i = 0; i < value.Count - 1; i++)
                                {
                                    if (value[i + 1].ToString() != "")
                                    {
                                        voteDirector.AddVoter(i, value[i + 1].ToString());
                                    }
                                }

                                if (value.Count == 1)
                                {
                                    break;
                                }
                            }

                            // 순위
                            index = 1;
                            for (int i = 4; index <= 8; index++)
                            {
                                value = values[index];

                                if (value[i + 1].ToString() != "")
                                {
                                    CVoteRanking ranking = new CVoteRanking();

                                    if ((value[i].ToString() == "1") || (value[i].ToString() == "2") || (value[i].ToString() == "3") || (value[i].ToString() == "4") ||
                                        (value[i].ToString() == "5") || (value[i].ToString() == "6") || (value[i].ToString() == "7") || (value[i].ToString() == "8"))
                                    {
                                        ranking.setRanking(Convert.ToInt32(value[i].ToString()), value[i + 1].ToString(), value[i + 2].ToString(), Convert.ToInt32(value[i + 3].ToString()), value[i + 4].ToString());
                                        voteDirector.AddRanking(ranking);
                                    }
                                    else
                                    {
                                        ranking.setRanking(0, value[i + 1].ToString(), value[i + 2].ToString(), Convert.ToInt32(value[i + 3].ToString()), value[i + 4].ToString());
                                        voteDirector.AddRanking(ranking);
                                    }
                                }
                            }

                            if (strContents == "")
                            {
                                strPrint += voteDirector.getVoteContents() + "\n";
                                strPrint += "=============================\n";
                                for (int i = 0; i < voteDirector.GetItemCount(); i++)
                                {
                                    strPrint += i + 1 + ". " + voteDirector.GetItem(i).getItem() + "\n";
                                }
                                strPrint += "\n \"/투표 숫자\"로 투표해주세요.";
                            }
                            else if (strContents == "결과")
                            {
                                strPrint += voteDirector.getVoteContents() + "\n";
                                strPrint += "=============================\n";
                                for (int i = 0; i < voteDirector.getRanking().Count; i++)
                                {
                                    var ranking = voteDirector.getRanking().ElementAt(i);
                                    strPrint += ranking.getRanking().ToString() + "위. " + ranking.getNumber() + " " + ranking.getVoteItem() + " [ " + ranking.getVoteCount() + "표 ] - " + ranking.getVoteRate() + "\n";
                                }
                            }
                            else
                            {
                                // 투표 했는지 체크
                                if (isAnonymous == false)
                                {
                                    // 익명 투표가 아닐 경우에만 시트로 체크
                                    for (int i = 0; i < voteDirector.GetItemCount(); i++)
                                    {
                                        List<string> voterCheckList = voteDirector.getVoter(i);

                                        foreach (var item in voterCheckList)
                                        {
                                            if (item == (strFirstName + strLastName))
                                            {
                                                await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 이미 투표를 하셨습니다.", ParseMode.Default, false, false, iMessageID);
                                                return;
                                            }
                                        }
                                    }
                                }
                                else // 익명 투표일 경우 파일에서 유저의 키로 중복 체크
                                {
                                    string[] voters = System.IO.File.ReadAllLines(@"_Voter.txt");
                                    foreach (string voter in voters)
                                    {
                                        if (voter.ToString() == senderKey.ToString())
                                        {
                                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 이미 투표를 하셨습니다.", ParseMode.Default, false, false, iMessageID);
                                            return;
                                        }
                                    }
                                }

                                string cellChar = "";

                                switch (strContents)
                                {
                                    case "1":
                                        cellChar = "C";
                                        break;
                                    case "2":
                                        cellChar = "D";
                                        break;
                                    case "3":
                                        cellChar = "E";
                                        break;
                                    case "4":
                                        cellChar = "F";
                                        break;
                                    case "5":
                                        cellChar = "G";
                                        break;
                                    case "6":
                                        cellChar = "H";
                                        break;
                                    case "7":
                                        cellChar = "I";
                                        break;
                                    case "8":
                                        cellChar = "J";
                                        break;
                                    default:
                                        {
                                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 투표 항목을 잘못 선택하셨습니다.", ParseMode.Default, false, false, iMessageID);
                                            return;
                                        }
                                }

                                int voteIndex = Convert.ToInt32(strContents);
                                if ((voteIndex <= 0) || (voteIndex > voteDirector.GetItemCount()))
                                {
                                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 투표 항목을 잘못 선택하셨습니다.", ParseMode.Default, false, false, iMessageID);
                                    return;
                                }

                                List<string> voterList = voteDirector.getVoter(voteIndex - 1);
                                int voterCount = voterList.Count;
                                string updateRange = "투표!" + cellChar + (18 + voterCount) + ":" + cellChar;

                                // Define request parameters.
                                SpreadsheetsResource.ValuesResource.GetRequest updateRequest = service.Spreadsheets.Values.Get(spreadsheetId, updateRange);
                                ValueRange valueRange = new ValueRange();
                                valueRange.MajorDimension = "COLUMNS"; //"ROWS";//COLUMNS 

                                string updateString = "";
                                if (isAnonymous == false)
                                {
                                    // 실명투표일 경우 대화명 입력
                                    updateString = strFirstName + strLastName;
                                }
                                else
                                {
                                    // 익명투표일 경우 O표시만
                                    updateString = "O";
                                }

                                var oblist = new List<object>() { updateString };
                                valueRange.Values = new List<IList<object>> { oblist };

                                SpreadsheetsResource.ValuesResource.UpdateRequest releaseRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, updateRange);

                                releaseRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                                UpdateValuesResponse releaseResponse = releaseRequest.Execute();
                                if (releaseResponse == null)
                                {
                                    strPrint = "[ERROR] 시트를 업데이트 할 수 없습니다.";
                                }
                                else
                                {
                                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[SUCCESS] 투표를 완료했습니다.", ParseMode.Default, false, false, iMessageID);
                                }

                                if (isAnonymous == true)
                                {
                                    System.IO.File.AppendAllText(@"_Voter.txt", senderKey.ToString() + "\n", Encoding.Default);
                                }

                                return;
                            }
                        }

                        if (strPrint != "")
                        {
                            const string vote = @"Function/Vote.jpg";
                            var fileName = vote.Split(Path.DirectorySeparatorChar).Last();
                            var fileStream = new FileStream(vote, FileMode.Open, FileAccess.Read, FileShare.Read);
                            await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream, strPrint, ParseMode.Default, false, iMessageID);
                        }
                        else
                        {
                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 현재 투표가 없습니다.", ParseMode.Default, false, false, iMessageID);
                        }
                    }
                }
            }
            //========================================================================================
            // 명예의 전당
            //========================================================================================
            else if (strCommend == "/기록")
            {
                if (strContents == "")
                {
                    // 내부 대회
                    String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                    String range = "명예의 전당!B7:F16";
                    SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                    ValueRange response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            int index = 1;
                            strPrint += "★ CDT 오버워치 리그 우승팀 ★\n==============================\n";

                            foreach (var row in values)
                            {
                                if (row.Count <= 0)
                                {
                                    break;
                                }

                                strPrint += "[ 1-" + index++ + " ] " + row[0].ToString() + " <" + row[1].ToString() + ">\n";
                            }
                        }
                    }

                    // 외부 대회
                    range = "명예의 전당!B21:F30";
                    request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                    response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            int index = 1;
                            strPrint += "\n★ 외부 대회 출전 ★\n==============================\n";

                            foreach (var row in values)
                            {
                                if (row.Count <= 0)
                                {
                                    break;
                                }

                                strPrint += "[ 2-" + index++ + " ] " + row[0].ToString() + " <" + row[1].ToString() + ">\n";
                            }
                        }
                    }

                    if (strPrint != "")
                    {
                        strPrint += "\n/기록 숫자 로 조회할 수 있습니다.\n(ex: /기록 2-3)";

                        const string record = @"Function/Record.jpg";
                        var fileName = record.Split(Path.DirectorySeparatorChar).Last();
                        var fileStream = new FileStream(record, FileMode.Open, FileAccess.Read, FileShare.Read);
                        await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream, strPrint, ParseMode.Default, false, iMessageID);
                    }
                    else
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 명예의 전당이 비어있습니다.", ParseMode.Default, false, false, iMessageID);
                    }
                }
                else
                {
                    string[] category = strContents.ToString().Split('-');
                    if (category.Count() <= 0)
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 항목을 잘못 입력했습니다.", ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    // 내부 or 외부
                    string upper = category[0].ToUpper();
                    if (upper.Length > 1)
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 항목을 잘못 입력했습니다.", ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    // 내부 대회
                    if (upper == "1")
                    {
                        // 항목
                        string strItem = category[1].ToString();
                        int item = Convert.ToInt32(strItem);
                        if (item < 1 || item > 999)
                        {
                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 항목을 잘못 입력했습니다.", ParseMode.Default, false, false, iMessageID);
                            return;
                        }

                        item--;

                        String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                        String range = "명예의 전당!B" + (7 + item) + ":F" + (7 + item);
                        SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                        ValueRange response = request.Execute();
                        if (response != null)
                        {
                            IList<IList<Object>> values = response.Values;
                            if (values != null && values.Count > 0)
                            {
                                strPrint += "★ CDT 오버워치 리그 우승팀 ★\n==============================\n";

                                var row = values[0];
                                string member = row[4].ToString().Replace("/", ",");

                                strPrint += "▷ " + row[0].ToString() + " 우승팀 [ " + row[2].ToString() + " ]\n";
                                strPrint += "* 팀장 : " + row[3].ToString() + "\n";
                                strPrint += "* 팀원 : " + member;
                            }
                        }
                    }
                    // 외부 대회
                    else if (upper == "2")
                    {
                        // 항목
                        string strItem = category[1].ToString();
                        int item = Convert.ToInt32(strItem);
                        if (item < 1 || item > 999)
                        {
                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 항목을 잘못 입력했습니다.", ParseMode.Default, false, false, iMessageID);
                            return;
                        }

                        item--;

                        String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                        String range = "명예의 전당!B" + (21 + item) + ":F" + (21 + item);
                        SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                        ValueRange response = request.Execute();
                        if (response != null)
                        {
                            IList<IList<Object>> values = response.Values;
                            if (values != null && values.Count > 0)
                            {
                                strPrint += "★ 외부 대회 출전 ★\n==============================\n";

                                var row = values[0];
                                string member = row[4].ToString().Replace("/", ",");

                                strPrint += "▷ " + row[0].ToString() + " [ " + row[2].ToString() + " ]\n";
                                strPrint += "* 팀장 : " + row[3].ToString() + "\n";
                                strPrint += "* 팀원 : " + member;
                            }
                        }
                    }

                    if (strPrint != "")
                    {
                        const string record = @"Function/Record.jpg";
                        var fileName = record.Split(Path.DirectorySeparatorChar).Last();
                        var fileStream = new FileStream(record, FileMode.Open, FileAccess.Read, FileShare.Read);
                        await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream, strPrint, ParseMode.Default, false, iMessageID);
                    }
                    else
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 명예의 전당이 비어있습니다.", ParseMode.Default, false, false, iMessageID);
                    }
                }
            }
            //========================================================================================
            // 스크림
            //========================================================================================
            else if ((strCommend == "/스크림") || (strCommend == "/디비전") || (strCommend == "/리그"))
            {
                string sheetName = "";
                int endLine = 0;

                if (strCommend == "/스크림")
                {
                    sheetName = "스크림";
                    endLine = 17;
                }
                else if (strCommend == "/디비전")
                {
                    sheetName = "디비전";
                    endLine = 17;
                }
                else if (strCommend == "/리그")
                {
                    sheetName = "리그 신청";
                    endLine = 35;
                }

                if (strContents == "")
                {
                    String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                    String range = sheetName + "!B2:U" + endLine;
                    SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                    ValueRange response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            // 스크림 이름
                            var title = values[0];
                            if (title.Count > 0 && title[0].ToString() != "")
                            {
                                strPrint += "[ " + title[0].ToString() + " ]\n============================\n";

                                int index = 4;
                                int totalCount = 0;
                                for (int i = index; i < values.Count; i++)
                                {
                                    var row = values[i];

                                    if (row.Count <= 1)
                                    {
                                        continue;
                                    }
                                    else if (row.Count <= 5)
                                    {
                                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 시트에 문제가 있습니다.\n시트를 확인해주세요.\n(ex: 빈칸 존재, 비정상 데이터 입력)", ParseMode.Default, false, false, iMessageID);
                                        return;
                                    }

                                    string battleTag = row[1].ToString();
                                    string tier = row[2].ToString();
                                    string score = row[3].ToString();
                                    string position = row[5].ToString();
                                    string date = "";
                                    if (row.Count > 13 && row[13].ToString() == "O")
                                        date += "월 ";
                                    if (row.Count > 14 && row[14].ToString() == "O")
                                        date += "화 ";
                                    if (row.Count > 15 && row[15].ToString() == "O")
                                        date += "수 ";
                                    if (row.Count > 16 && row[16].ToString() == "O")
                                        date += "목 ";
                                    if (row.Count > 17 && row[17].ToString() == "O")
                                        date += "금 ";
                                    if (row.Count > 18 && row[18].ToString() == "O")
                                        date += "토 ";
                                    if (row.Count > 19 && row[19].ToString() == "O")
                                        date += "일";

                                    strPrint += "- " + battleTag.ToString() + " (" + position.ToString() + ") / " + score.ToString() + " - " + date.ToString() + "\n";
                                    totalCount++;
                                }

                                strPrint += "\n현재 " + totalCount + "명\n";
                            }
                        }
                    }

                    if (strPrint != "")
                    {
                        strPrint += "\n신청은 /" + sheetName + " [요일] 로 해주세요.\n(ex: /" + sheetName + " 토일)";

                        string scrim = "";

                        if ((sheetName == "스크림") || (sheetName == "리그 신청"))
                            scrim = @"Function/Scrim.png";
                        else if (sheetName == "디비전")
                            scrim = @"Function/OpenDivision.png";
                        else
                            return;

                        var fileName = scrim.Split(Path.DirectorySeparatorChar).Last();
                        var fileStream = new FileStream(scrim, FileMode.Open, FileAccess.Read, FileShare.Read);
                        await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream, strPrint, ParseMode.Default, false, iMessageID);
                    }
                    else
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 현재 모집 중인 " + sheetName + "이 없습니다.", ParseMode.Default, false, false, iMessageID);
                    }
                }
                else // 일정을 입력했을 경우
                {
                    // 타이틀 Load
                    {
                        String sheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                        String titleRange = sheetName + "!B2:B";
                        SpreadsheetsResource.ValuesResource.GetRequest titleRequest = service.Spreadsheets.Values.Get(sheetId, titleRange);

                        ValueRange titleResponse = titleRequest.Execute();
                        if (titleResponse != null)
                        {
                            IList<IList<Object>> values = titleResponse.Values;
                            if (values != null && values.Count > 0)
                            {
                                // 스크림 이름
                                var title = values[0];
                                if (title.Count == 0 || title[0].ToString() == "")
                                {
                                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 현재 모집 중인 " + sheetName + "이 없습니다.", ParseMode.Default, false, false, iMessageID);
                                    return;
                                }
                            }
                        }
                    }

                    int size = strContents.Length;
                    string[] day = { "", "", "", "", "", "", "" };
                    bool isConfirmDay = false;
                    bool isCancel = false;

                    if (strContents == "취소")
                    {
                        isCancel = true;
                    }
                    else
                    {
                        if (strContents.Contains("요") == true)
                        {
                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 잘못된 날짜입니다.", ParseMode.Default, false, false, iMessageID);
                            return;
                        }

                        // 가능 날짜 추출
                        for (int i = 0; i < size; i++)
                        {
                            string inputDay = strContents.Substring(i, 1);

                            if (inputDay == "월")
                            {
                                day[0] = "O";
                                isConfirmDay = true;
                            }
                            if (inputDay == "화")
                            {
                                day[1] = "O";
                                isConfirmDay = true;
                            }
                            if (inputDay == "수")
                            {
                                day[2] = "O";
                                isConfirmDay = true;
                            }
                            if (inputDay == "목")
                            {
                                day[3] = "O";
                                isConfirmDay = true;
                            }
                            if (inputDay == "금")
                            {
                                day[4] = "O";
                                isConfirmDay = true;
                            }
                            if (inputDay == "토")
                            {
                                day[5] = "O";
                                isConfirmDay = true;
                            }
                            if (inputDay == "일")
                            {
                                day[6] = "O";
                                isConfirmDay = true;
                            }
                        }

                        if (isConfirmDay == false)
                        {
                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 잘못된 날짜입니다.", ParseMode.Default, false, false, iMessageID);
                            return;
                        }
                    }

                    String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                    CUser user = new CUser();

                    // 클랜원 목록에서 정보 추출
                    String range = "클랜원 목록!C8:V";
                    SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                    ValueRange response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            foreach (var row in values)
                            {
                                if (row[18] == null || row[18].ToString() == "")
                                    continue;

                                // 유저키 일치
                                if (Convert.ToInt64(row[18].ToString()) == senderKey)
                                {
                                    user = setUserInfo(row, senderKey);
                                    break;
                                }
                            }
                        }
                    }

                    int index = 0;
                    bool isInput = false;

                    range = sheetName + "!C6:C" + endLine;
                    request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                    response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            if ((sheetName == "스크림" && values.Count >= 12) ||
                                (sheetName == "디비전" && values.Count >= 12) ||
                                (sheetName == "리그 신청" && values.Count >= 30))
                            {
                                await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[SYSTEM] " + sheetName + " 신청자가 모두 찼습니다.\n운영진에게 문의해주세요.", ParseMode.Default, false, false, iMessageID);
                                return;
                            }

                            int count = 0;
                            bool isSearch = false;

                            foreach (var row in values)
                            {
                                if (row.Count > 0 && row[0].ToString() != "")
                                {
                                    if (row[0].ToString() == user.MainBattleTag)
                                    {
                                        isSearch = true;
                                        index = count;
                                        break;
                                    }

                                    count++;
                                }
                                else
                                {
                                    if (isInput == false)
                                    {
                                        index = count;
                                        isInput = true;
                                    }
                                }
                            }

                            if (isInput == false)
                            {
                                index = count;
                            }

                            if (isCancel == true && isSearch == false)
                            {
                                await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] " + sheetName + " 신청을 하지 않았습니다.", ParseMode.Default, false, false, iMessageID);
                                return;
                            }
                        }
                    }

                    // 유저키 등록이 되어있다면
                    if (user.UserKey > 0)
                    {
                        // 스크림 신청
                        if (isCancel == false)
                        {
                            string position = "";

                            if (user.Position.HasFlag(POSITION.POSITION_FLEX) == true)
                            {
                                position = "플렉스";
                            }
                            else
                            {
                                if (user.Position.HasFlag(POSITION.POSITION_DPS) == true)
                                    position += "딜";
                                if (user.Position.HasFlag(POSITION.POSITION_TANK) == true)
                                    position += "탱";
                                if (user.Position.HasFlag(POSITION.POSITION_SUPP) == true)
                                    position += "힐";
                            }

                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, "'" + user.MainBattleTag + "'의 전적을 조회 중입니다.\n잠시만 기다려주세요.", ParseMode.Default, false, false, iMessageID);

                            Tuple<int, string> retTuple = referenceScore(user.MainBattleTag);
                            int score = retTuple.Item1;     // 점수
                            string tier = retTuple.Item2;   // 티어

                            if (score == 0)
                            {
                                await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 전적을 조회 할 수 없습니다.\n시트에 직접 입력해주세요.\n(원인 : 프로필 비공개, 친구공개, 미배치)", ParseMode.Default, false, false, iMessageID);
                            }

                            // Define request parameters.
                            range = sheetName + "!C" + (6 + index) + ":U" + (6 + index);
                            ValueRange valueRange = new ValueRange();
                            valueRange.MajorDimension = "ROWS"; //"ROWS";//COLUMNS 

                            var oblist = new List<object>()
                            {
                                user.MainBattleTag, // 배틀태그
                                tier,               // 티어
                                score,              // 점수
                                "",                 // 6명 체크
                                position,           // 포지션
                                user.MostPick[0],   // 모스트1
                                user.MostPick[1],   // 모스트2
                                user.MostPick[2],   // 모스트3
                                user.OtherPick,     // 이외 가능 픽
                                "",
                                "",
                                "",
                                day[0],             // 월
                                day[1],             // 화
                                day[2],             // 수
                                day[3],             // 목
                                day[4],             // 금
                                day[5],             // 토
                                day[6]              // 일
                            };
                            valueRange.Values = new List<IList<object>> { oblist };

                            SpreadsheetsResource.ValuesResource.UpdateRequest updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);

                            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                            UpdateValuesResponse updateResponse = updateRequest.Execute();
                            if (updateResponse == null)
                            {
                                strPrint += "[ERROR] 시트를 업데이트 할 수 없습니다.";
                            }
                            else
                            {
                                strPrint += "[SYSTEM] " + sheetName + " 신청이 완료 됐습니다.";
                            }
                        }
                        else
                        {
                            // 스크림 취소
                            // Define request parameters.
                            range = sheetName + "!C" + (6 + index) + ":Z" + (6 + index);
                            ValueRange valueRange = new ValueRange();
                            valueRange.MajorDimension = "ROWS"; //"ROWS";//COLUMNS 

                            var oblist = new List<object>() { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
                            valueRange.Values = new List<IList<object>> { oblist };

                            SpreadsheetsResource.ValuesResource.UpdateRequest updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);

                            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                            UpdateValuesResponse updateResponse = updateRequest.Execute();
                            if (updateResponse == null)
                            {
                                strPrint += "[ERROR] 시트를 업데이트 할 수 없습니다.";
                            }
                            else
                            {
                                strPrint += "[SYSTEM] " + sheetName + " 신청을 취소했습니다.";
                            }
                        }
                    }
                    else
                    {
                        strPrint += "[ERROR] 유저 정보를 업데이트 할 수 없습니다.";
                    }

                    if (strPrint != "")
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                    }
                }
            }
            //========================================================================================
            // 일정 조사
            //========================================================================================
            else if (strCommend == "/조사")
            {
                if (strContents == "")
                {
                    String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                    String range = "일정 조사!L7:R14";
                    SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                    ValueRange response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            var title = values[0];
                            if (title.Count == 0)
                            {
                                strPrint += "[ERROR] 현재 조사 중인 일정이 없습니다.";
                                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                                return;
                            }
                            else
                            {
                                strPrint += title[0].ToString() + "\n============================\n";

                                var day = values[6];
                                var count = values[7];

                                for (int i = 0; i < 7; i++)
                                {
                                    if (count[i].ToString() != "0")
                                    {
                                        strPrint += "- " + day[i].ToString() + " : " + count[i].ToString() + "명\n";
                                    }
                                }
                            }
                        }
                    }

                    if (strPrint != "")
                    {
                        strPrint += "\n조사에 참여하려면 /조사 [요일] 로 참여해주세요.\n(ex: /조사 금토일)";

                        const string calendar_research = @"Function/calendar_research.jpg";
                        var fileName = calendar_research.Split(Path.DirectorySeparatorChar).Last();
                        var fileStream = new FileStream(calendar_research, FileMode.Open, FileAccess.Read, FileShare.Read);
                        await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream, strPrint, ParseMode.Default, false, iMessageID);
                    }
                    else
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 현재 모집 중인 스크림이 없습니다.", ParseMode.Default, false, false, iMessageID);
                    }
                }
                else
                {
                    int size = strContents.Length;
                    string[] day = { "", "", "", "", "", "", "" };
                    bool isConfirmDay = false;
                    bool isCancel = false;

                    if (strContents == "취소")
                    {
                        isCancel = true;
                    }
                    else
                    {
                        if (strContents.Contains("요") == true)
                        {
                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 잘못된 날짜입니다.", ParseMode.Default, false, false, iMessageID);
                            return;
                        }

                        // 가능 날짜 추출
                        for (int i = 0; i < size; i++)
                        {
                            string inputDay = strContents.Substring(i, 1);

                            if (inputDay == "월")
                            {
                                day[0] = "O";
                                isConfirmDay = true;
                            }
                            if (inputDay == "화")
                            {
                                day[1] = "O";
                                isConfirmDay = true;
                            }
                            if (inputDay == "수")
                            {
                                day[2] = "O";
                                isConfirmDay = true;
                            }
                            if (inputDay == "목")
                            {
                                day[3] = "O";
                                isConfirmDay = true;
                            }
                            if (inputDay == "금")
                            {
                                day[4] = "O";
                                isConfirmDay = true;
                            }
                            if (inputDay == "토")
                            {
                                day[5] = "O";
                                isConfirmDay = true;
                            }
                            if (inputDay == "일")
                            {
                                day[6] = "O";
                                isConfirmDay = true;
                            }
                        }

                        if (isConfirmDay == false)
                        {
                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 잘못된 날짜입니다.", ParseMode.Default, false, false, iMessageID);
                            return;
                        }
                    }

                    string calTitle = "";

                    String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                    String range = "일정 조사!L7:L";
                    SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                    ValueRange response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            var title = values[0];
                            if (title.Count == 0)
                            {
                                strPrint += "[ERROR] 현재 조사 중인 일정이 없습니다.";
                                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                                return;
                            }
                            else
                            {
                                // 일정 조사 제목
                                calTitle = title[0].ToString();
                            }
                        }
                    }

                    int index = 0;

                    spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                    range = "일정 조사!B5:J74";
                    request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                    response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            int count = 0;
                            bool isInput = false;
                            bool isBlank = false;

                            foreach (var row in values)
                            {
                                if (row.Count > 1 && row[1].ToString() != "")
                                {
                                    if (row[1].ToString() == strUserName)
                                    {
                                        index = count;
                                        isInput = true;
                                        break;
                                    }
                                    else
                                    {
                                        count++;
                                    }
                                }
                                else
                                {
                                    if (index == 0 && isBlank == false)
                                    {
                                        index = count++;
                                        isBlank = true;
                                    }
                                }
                            }

                            if (isInput == false && isBlank == false && index == 0)
                            {
                                index = count;
                            }
                        }
                    }

                    if (isCancel == false)
                    {
                        range = "일정 조사!C" + (5 + index) + ":J" + (5 + index);
                        ValueRange valueRange = new ValueRange();
                        valueRange.MajorDimension = "ROWS"; //"ROWS";//COLUMNS 

                        var oblist = new List<object>()
                            {
                                strUserName,
                                day[0],             // 월
                                day[1],             // 화
                                day[2],             // 수
                                day[3],             // 목
                                day[4],             // 금
                                day[5],             // 토
                                day[6]              // 일
                            };
                        valueRange.Values = new List<IList<object>> { oblist };

                        SpreadsheetsResource.ValuesResource.UpdateRequest updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);

                        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                        UpdateValuesResponse updateResponse = updateRequest.Execute();
                        if (updateResponse == null)
                        {
                            strPrint += "[ERROR] 시트를 업데이트 할 수 없습니다.";
                        }
                        else
                        {
                            strPrint += "[SYSTEM] 일정 조사를 완료했습니다.";
                        }
                    }
                    else
                    {
                        range = "일정 조사!C" + (5 + index) + ":J" + (5 + index);
                        ValueRange valueRange = new ValueRange();
                        valueRange.MajorDimension = "ROWS"; //"ROWS";//COLUMNS 

                        var oblist = new List<object>() { "", "", "", "", "", "", "", "" };
                        valueRange.Values = new List<IList<object>> { oblist };

                        SpreadsheetsResource.ValuesResource.UpdateRequest updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);

                        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                        UpdateValuesResponse updateResponse = updateRequest.Execute();
                        if (updateResponse == null)
                        {
                            strPrint += "[ERROR] 시트를 업데이트 할 수 없습니다.";
                        }
                        else
                        {
                            strPrint += "[SYSTEM] 일정 조사를 취소했습니다.";
                        }
                    }

                    if (strPrint != "")
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                    }
                    else
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 일정 조사를 할 수 없습니다.", ParseMode.Default, false, false, iMessageID);
                    }
                }
            }
            //========================================================================================
            // NAS 파일 리스트
            //========================================================================================
            else if (strCommend == "/dir")
            {
                // 관리자 전용 명령어
                if (varMessage.Chat.Id != config.getGroupKey(GROUP_TYPE.GROUP_TYPE_TEST))
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 사용할 수 없는 대화방입니다.", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                if (userDirector.getUserInfo(senderKey).UserType != USER_TYPE.USER_TYPE_ADMIN &&
                    config.isDeveloper(senderKey) == false)
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 권한이 없는 유저 또는 대화방입니다.", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                String FolderName = nasInfo.CurrentPath;
                System.IO.DirectoryInfo directory = new System.IO.DirectoryInfo(FolderName);
                if (directory.Exists == true)
                {
                    strPrint += nasInfo.CurrentPath + "\n\n";

                    foreach (var file in directory.GetDirectories())
                    {
                        strPrint += "[D] " + file.Name + "\n";
                    }
                    foreach (var file in directory.GetFiles())
                    {
                        strPrint += file.Name + "\n";
                    }
                }

                if (strPrint != "")
                {
                    nasInfo.CurrentPath = FolderName;
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[SYSTEM] 경로가 존재하지 않습니다.\n" + FolderName, ParseMode.Default, false, false, iMessageID);
                }
            }
            //========================================================================================
            // NAS 폴더 이동
            //========================================================================================
            else if (strCommend == "/cd")
            {
                // 관리자 전용 명령어
                if (varMessage.Chat.Id != config.getGroupKey(GROUP_TYPE.GROUP_TYPE_TEST))
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 사용할 수 없는 대화방입니다.", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                if (userDirector.getUserInfo(senderKey).UserType != USER_TYPE.USER_TYPE_ADMIN &&
                    config.isDeveloper(senderKey) == false)
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 권한이 없는 유저 또는 대화방입니다.", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                // 현재 경로
                if (strContents == "")
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, nasInfo.CurrentPath, ParseMode.Default, false, false, iMessageID);
                    return;
                }
                // 한 단계 위로 이동
                else if (strContents == "..")
                {
                    if (nasInfo.CurrentPath == @"D:\CDT\")
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "최상위 경로입니다.", ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    nasInfo.CurrentPath = nasInfo.CurrentPath.Substring(0, nasInfo.CurrentPath.LastIndexOf('\\'));
                    nasInfo.CurrentPath = nasInfo.CurrentPath.Substring(0, nasInfo.CurrentPath.LastIndexOf('\\'));
                    nasInfo.CurrentPath += @"\";
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, nasInfo.CurrentPath, ParseMode.Default, false, false, iMessageID);
                }
                // 최상위 경로로 이동
                else if (strContents == "\\")
                {
                    nasInfo.CurrentPath = @"D:\CDT\";
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, nasInfo.CurrentPath, ParseMode.Default, false, false, iMessageID);
                    return;
                }
                else
                {
                    String FolderName = nasInfo.CurrentPath + strContents + @"\";
                    System.IO.DirectoryInfo directory = new System.IO.DirectoryInfo(FolderName);
                    if (directory.Exists == true)
                    {
                        nasInfo.CurrentPath = FolderName;
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, nasInfo.CurrentPath, ParseMode.Default, false, false, iMessageID);
                    }
                    else
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[SYSTEM] 경로가 존재하지 않습니다.\n" + FolderName, ParseMode.Default, false, false, iMessageID);
                    }
                }
            }
            //========================================================================================
            // NAS 파일 다운로드
            //========================================================================================
            else if (strCommend == "/down")
            {
                // 관리자 전용 명령어
                if (varMessage.Chat.Id != config.getGroupKey(GROUP_TYPE.GROUP_TYPE_TEST))
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 사용할 수 없는 대화방입니다.", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                if (userDirector.getUserInfo(senderKey).UserType != USER_TYPE.USER_TYPE_ADMIN &&
                    config.isDeveloper(senderKey) == false)
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 권한이 없는 유저 또는 대화방입니다.", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                if (strContents == "")
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 파일명을 입력해주세요.", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                string FileName = nasInfo.CurrentPath + strContents;
                System.IO.FileInfo file = new System.IO.FileInfo(FileName);
                if (file.Exists == true)
                {
                    var fileStream = new FileStream(FileName, FileMode.Open, FileAccess.Read, FileShare.Read);
                    var stream = new Telegram.Bot.Types.InputFiles.InputOnlineFile(fileStream, strContents);

                    await Bot.SendDocumentAsync(varMessage.Chat.Id, stream);
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[SYSTEM] 파일이 존재하지 않습니다.\n" + FileName, ParseMode.Default, false, false, iMessageID);
                }
            }
            //========================================================================================
            // NAS 파일 다운로드
            //========================================================================================
            else if (strCommend == "/up")
            {
                // 관리자 전용 명령어
                if (varMessage.Chat.Id != config.getGroupKey(GROUP_TYPE.GROUP_TYPE_TEST))
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 사용할 수 없는 대화방입니다.", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                if (userDirector.getUserInfo(senderKey).UserType != USER_TYPE.USER_TYPE_ADMIN &&
                    config.isDeveloper(senderKey) == false)
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 권한이 없는 유저 또는 대화방입니다.", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                // Download File
                var file = await Bot.GetFileAsync(varMessage.Document.FileId);
                var fileName = nasInfo.CurrentPath + varMessage.Document.FileName;
                string fileExtention = fileName.Substring(fileName.LastIndexOf('.')).Replace(".", "");
                string onlyFileName = varMessage.Document.FileName.Substring(0, varMessage.Document.FileName.LastIndexOf('.'));
                int fileNameIndex = 2;
                string fileStatue = "";

                while (true)
                {
                    System.IO.FileInfo fileInfo = new System.IO.FileInfo(fileName);
                    if (fileInfo.Exists == true)
                    {
                        fileName = nasInfo.CurrentPath + onlyFileName + " (" + fileNameIndex++ + ")." + fileExtention;
                        fileStatue = "\n(같은 이름을 가진 파일이 있어서 파일명을 변경합니다.)";
                    }
                    else
                    {
                        break;
                    }
                }

                using (var saveStream = new FileStream(fileName, FileMode.CreateNew, FileAccess.Write))
                {
                    await Bot.DownloadFileAsync(file.FilePath, saveStream);
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[SYSTEM] 파일 업로드 완료.\n" + fileName + fileStatue, ParseMode.Default, false, false, iMessageID);
                }
            }
            //========================================================================================
            // 뽑기
            //========================================================================================
            else if (strCommend == "/뽑기")
            {
                if (strContents == "")
                {
                    strPrint += "[SYSTEM] 뽑을 항목을 추가해주세요.\n (ex: /뽑기 [항목1] [항목2] ...)";
                }
                else
                {
                    string[] item = strContents.Split(' ');

                    Random random = new Random();
                    int index = random.Next(0, item.Count());

                    strPrint += item.ElementAt(index);
                }

                if (strPrint != "")
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 뽑기를 할 수 없습니다.", ParseMode.Default, false, false, iMessageID);
                }
            }
            //========================================================================================
            // 날씨
            //========================================================================================
            else if (strCommend == "/날씨")
            {
                if (strContents == "")
                {
                    strPrint += "[SYSTEM] 지역을 추가해주세요.\n- 가능지역 : 서울, 경기, 부산, 대구, 광주, 인천, 대전, 울산, 세종, 제주\n(ex: /날씨 제주)";
                }
                else
                {
                    Tuple<string, string> city = CWeather.getCity(strContents);

                    if (city.Item1 == "" || city.Item2 == "")
                    {
                        strPrint += "[ERROR] 지역을 다시 확인해주세요.\n";
                    }
                    else
                    {
                        // OpenWeatherMap 날씨
                        string weatherUrl = "http://api.openweathermap.org/data/2.5/weather?q=" + city.Item1.ToString() + ",KR&appid=604e54fc977920798ff275b8da0a687f";

                        try
                        {
                            WebClient wc = new WebClient();
                            wc.Encoding = Encoding.UTF8;

                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

                            string html = wc.DownloadString(weatherUrl);

                            var json = JObject.Parse(html);

                            string main = json["weather"][0]["main"].ToString();
                            string temp = Math.Round(Convert.ToDouble(json["main"]["temp"].ToString()) - 273.15, 1).ToString();
                            string humidity = json["main"]["humidity"].ToString();

                            strPrint += "[ " + city.Item2 + "의 날씨 ]\n============================\n";
                            strPrint += "- 날씨요약 : " + main + "\n";
                            strPrint += "- 현재기온 : " + temp + "℃\n";
                            strPrint += "- 현재습도 : " + humidity + "%\n";
                        }
                        catch
                        {
                            // 없을 수도 있다.
                            //await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 날씨를 조회할 수 없습니다.", ParseMode.Default, false, false, iMessageID);
                            //return;
                        }

                        // 공공데이터포럼 미세먼지
                        weatherUrl = "http://openapi.airkorea.or.kr/openapi/services/rest/ArpltnInforInqireSvc/getCtprvnRltmMesureDnsty?sidoName=" + city.Item2.ToString() + "&pageNo=1&numOfRows=350&ServiceKey=wg%2FwOk%2Fmolt2CQfeZPTss%2BkroS0o0ygHBPR%2BXoGjPEEJyUYYhvMv1mi1D0kWsSjSEQy7ctTH4sZA1IV2816U8Q%3D%3D&ver=1.3&_returnType=json";

                        try
                        {
                            WebClient wc = new WebClient();
                            wc.Encoding = Encoding.UTF8;

                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

                            string html = wc.DownloadString(weatherUrl);

                            var json = JObject.Parse(html);
                            uint count = Convert.ToUInt32(json["totalCount"].ToString());
                            uint pm10Value = 0;
                            uint pm25Value = 0;
                            uint pm10Grade = 0;
                            uint pm25Grade = 0;
                            uint pm10ValueCount = 0;
                            uint pm10GradeCount = 0;
                            uint pm25ValueCount = 0;
                            uint pm25GradeCount = 0;

                            for (int i = 0; i < count; i++)
                            {
                                if (json["list"][i]["pm10Grade"].ToString() != "" && json["list"][i]["pm10Grade"].ToString() != "-")
                                {
                                    pm10Grade += Convert.ToUInt32(json["list"][i]["pm10Grade"].ToString());
                                    pm10GradeCount++;
                                }

                                if (json["list"][i]["pm10Value"].ToString() != "" && json["list"][i]["pm10Value"].ToString() != "-")
                                {
                                    pm10Value += Convert.ToUInt32(json["list"][i]["pm10Value"].ToString());
                                    pm10ValueCount++;
                                }

                                if (json["list"][i]["pm25Grade"].ToString() != "" && json["list"][i]["pm25Grade"].ToString() != "-")
                                {
                                    pm25Grade += Convert.ToUInt32(json["list"][i]["pm25Grade"].ToString());
                                    pm25GradeCount++;
                                }

                                if (json["list"][i]["pm25Value"].ToString() != "" && json["list"][i]["pm25Value"].ToString() != "-")
                                {
                                    pm25Value += Convert.ToUInt32(json["list"][i]["pm25Value"].ToString());
                                    pm25ValueCount++;
                                }
                            }

                            pm10Value = Convert.ToUInt32(pm10Value / pm10ValueCount);
                            pm25Value = Convert.ToUInt32(pm25Value / pm10GradeCount);
                            pm10Grade = Convert.ToUInt32(pm10Grade / pm25GradeCount);
                            pm25Grade = Convert.ToUInt32(pm25Grade / pm25ValueCount);

                            string pm10GradeString = "";
                            string pm25GradeString = "";

                            switch (pm10Grade)
                            {
                                case 1:
                                    pm10GradeString = "좋음";
                                    break;
                                case 2:
                                    pm10GradeString = "보통";
                                    break;
                                case 3:
                                    pm10GradeString = "나쁨";
                                    break;
                                case 4:
                                    pm10GradeString = "매우나쁨";
                                    break;
                            }

                            switch (pm25Grade)
                            {
                                case 1:
                                    pm25GradeString = "좋음";
                                    break;
                                case 2:
                                    pm25GradeString = "보통";
                                    break;
                                case 3:
                                    pm25GradeString = "나쁨";
                                    break;
                                case 4:
                                    pm25GradeString = "매우나쁨";
                                    break;
                            }

                            strPrint += "- 미세먼지 : " + pm10GradeString.ToString() + "(" + pm10Value.ToString() + ")\n";
                            strPrint += "- 초미세먼지 : " + pm25GradeString.ToString() + "(" + pm25Value.ToString() + ")\n";
                        }
                        catch
                        {
                            // 미세먼지는 지역이 없을 수 있다.
                            //await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 날씨를 조회할 수 없습니다.", ParseMode.Default, false, false, iMessageID);
                            //return;
                        }
                    }
                }

                if (strPrint != "")
                {
                    strPrint += "\n* 자료제공\n(날씨) OpenWeatherMap\n(대기) 한국환경공단, 공공데이터포럼";
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 날씨를 조회할 수 없습니다.", ParseMode.Default, false, false, iMessageID);
                }
            }
            //========================================================================================
            // 알림
            //========================================================================================
            else if (strCommend == "/알림")
            {
                if (strUserID == null || strUserID == "")
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 텔레그램 아이디를 먼저 설정해주세요.", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                if (strContents == "")
                {
                    if (userDirector.getPrivateNoti(senderKey).Count == 0)
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 설정된 알림이 없습니다.", ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    strPrint += "[ " + strUserName + "님의 알림 ]\n";
                    strPrint += "--------------------------------------";

                    foreach (var elem in userDirector.getPrivateNoti(senderKey))
                    {
                        strPrint += "\n(" + elem.Hour + "시 " + elem.Minute + "분) " + elem.Notice;
                    }

                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                    return;
                }
                else if (strContents.Length >= 2 && strContents.Substring(0, 2) == "제거")
                {
                    int hour = Convert.ToInt32(strContents.Substring(3, 2));
                    int min = Convert.ToInt32(strContents.Substring(5, 2));
                    int index = 0;

                    foreach (var elem in userDirector.getPrivateNoti(senderKey))
                    {
                        if (elem.Hour == hour && elem.Minute == min)
                        {
                            userDirector.RemoveNoti(senderKey, index);
                            strPrint += "[SYSTEM] 해당 알림이 제거 되었습니다.";
                            break;
                        }

                        index++;
                    }
                }
                else
                {
                    int checkNum = 0;
                    if (strContents.Length < 4)
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 시간을 잘못 입력하셨습니다.\n(ex: 오후 7시 30분 : 1930)", ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    bool isNumber = int.TryParse(strContents.Substring(0, 4), out checkNum);
                    if (isNumber == false)
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 시간을 잘못 입력하셨습니다.\n(ex: 오후 7시 30분 : 1930)", ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    bool isSearch = false;

                    string notiTime = strContents.Substring(0, 4);

                    if (strContents.Length < 5)
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 내용을 잘못 입력하셨습니다.", ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    string notiString = strContents.Substring(5);
                    if (notiString == "")
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 내용을 잘못 입력하셨습니다.", ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    int hour = Convert.ToInt32(notiTime.Substring(0, 2));
                    int min = Convert.ToInt32(notiTime.Substring(2, 2));

                    if (hour > 23 || min > 59)
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 시간을 잘못 입력하셨습니다.\n(ex: 오후 7시 30분 : 1930)", ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    if (hour != 0)
                    {
                        foreach (var elem in userDirector.getPrivateNoti(senderKey))
                        {
                            if (elem.Hour == hour && elem.Minute == min)
                            {
                                strPrint += "[ERROR] 해당 시간에 이미 알림이 있습니다.";
                                isSearch = true;
                                break;
                            }
                        }

                        if (isSearch == false)
                        {
                            userDirector.addPrivateNoti(senderKey, strUserID, notiString, hour, min);
                            strPrint += "[SYSTEM] 알림이 적용 되었습니다.";
                        }
                    }
                }

                if (strPrint != "")
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 알림을 적용할 수 없습니다.", ParseMode.Default, false, false, iMessageID);
                }
            }
            //========================================================================================
            // 메모
            //========================================================================================
            else if (strCommend == "/메모")
            {
                if (strContents == "")
                {
                    if (userDirector.getMemo(senderKey).Count == 0)
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 저장된 메모가 없습니다.", ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    strPrint += "[ " + strUserName + "님의 메모 ]\n";
                    strPrint += "--------------------------------------";

                    int index = 1;
                    foreach (var elem in userDirector.getMemo(senderKey))
                    {
                        strPrint += "\n(" + index++ + ") " + elem.Memo;
                    }

                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                    return;
                }
                else
                {
                    string[] contents = strContents.Split(' ');

                    if (contents[0] == "제거" && contents[1] != "")
                    {
                        int index = Convert.ToInt32(contents[1]);

                        if (userDirector.RemoveMemo(senderKey, index - 1) == true)
                            strPrint += "[SYSTEM] 해당 메모가 제거 되었습니다.";
                        else
                            strPrint += "[ERROR] 메모 인덱스가 잘못 됐습니다.";
                    }
                    else if (contents[0] != "제거" && contents[0] != "")
                    {
                        if (userDirector.getUserInfo(senderKey).getMemoList().Count >= 15)
                        {
                            userDirector.getUserInfo(senderKey).getMemoList().RemoveAt(0);
                            //userDirector.RemoveMemo(senderKey, 0);
                            userDirector.addMemo(senderKey, strContents);

                            strPrint += "[SYSTEM] 메모 제한 갯수를 초과하여\n가장 오래된 메모가 제거 되었으며,\n해당 메모가 저장 되었습니다.";
                        }
                        else
                        {
                            userDirector.addMemo(senderKey, strContents);
                            strPrint += "[SYSTEM] 해당 메모가 저장 되었습니다.";
                        }
                    }
                }

                if (strPrint != "")
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 메모를 적용할 수 없습니다.", ParseMode.Default, false, false, iMessageID);
                }
            }
            //========================================================================================
            // 문의
            //========================================================================================
            //else if (strCommend == "/문의")
            //{
            //    if (strContents == "")
            //    {
            //        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 문의 내용을 입력해주세요.", ParseMode.Default, false, false, iMessageID);
            //        return;
            //    }

            //    string name = userDirector.getUserInfo(senderKey).Name;
            //    strPrint += "[ " + name + "님의 문의 ]\n\n";
            //    strPrint += strContents + "\n\n";
            //    strPrint += "@Seungman / @hyulin / @mans3ul / @urusaikara / @jandie99";

            //    await Bot.SendTextMessageAsync(-1001482490165, strPrint);
            //    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[SYSTEM] 문의 등록이 완료 됐습니다.", ParseMode.Default, false, false, iMessageID);
            //}
            //========================================================================================
            // 차단
            //========================================================================================
            else if (strCommend == "/차단")
            {
                var user = userDirector.getUserInfo(senderKey);
                if (user.UserType != USER_TYPE.USER_TYPE_ADMIN && config.isDeveloper(senderKey) == false)
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 해당 명렁어는 관리자 권한이 필요합니다.", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                long blockUserKey = 0;
                bool bIsUser = false;

                if (strContents == "")
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 차단할 유저를 입력해주세요.", ParseMode.Default, false, false, iMessageID);
                    return;
                }
                else
                {
                    // Define request parameters.
                    String spreadsheetId = config.getTokenKey(TOKEN_TYPE.TOKEN_TYPE_SHEET);
                    String range = "클랜원 목록!C8:V";
                    SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                    ValueRange response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            foreach (var row in values)
                            {
                                if (row[0].ToString().Contains(strContents) == true)
                                {
                                    blockUserKey = Convert.ToInt64(row[18].ToString());
                                    if (userDirector.getUserInfo(blockUserKey).UserType != USER_TYPE.USER_TYPE_ADMIN &&
                                        config.isDeveloper(blockUserKey) == false)
                                    {
                                        if (bIsUser == true)
                                        {
                                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 검색된 유저가 다수입니다.", ParseMode.Default, false, false, iMessageID);
                                            return;
                                        }

                                        bIsUser = true;
                                    }
                                    else
                                    {
                                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 관리자는 차단할 수 없습니다.", ParseMode.Default, false, false, iMessageID);
                                        return;
                                    }
                                }
                            }
                        }
                    }

                    string[] contents = strContents.Split(' ');
                    if (contents[0] == "해제")
                    {
                        userDirector.removeBlockUser(blockUserKey);

                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[SYSTEM] 아테나 사용 차단이 해제되었습니다.", ParseMode.Default, false, false, iMessageID);
                        return;
                    }
                }

                if (bIsUser == true && blockUserKey != 0)
                {
                    userDirector.addBlockUser(blockUserKey);

                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[SYSTEM] 아테나 사용 차단되었습니다.", ParseMode.Default, false, false, iMessageID);
                    return;
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 차단할 유저가 없습니다.", ParseMode.Default, false, false, iMessageID);
                    return;
                }
            }
            //========================================================================================
            // 방등록
            //========================================================================================
            else if (strCommend == "/방등록")
            {
                var user = userDirector.getUserInfo(senderKey);
                if (user.UserType != USER_TYPE.USER_TYPE_ADMIN && config.isDeveloper(senderKey) == false)
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 해당 명렁어는 관리자 권한이 필요합니다.", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                if (strContents == "")
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 등록할 방을 입력해주세요.", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                if (strContents == "운영")
                {

                }
                else if (strContents == "클랜")
                {

                }
                else if (strContents == "안내")
                {

                }
            }
            //========================================================================================
            // 대화량 측정 리셋
            //========================================================================================
            else if (strCommend == "/리셋")
            {
                if (config.isDeveloper(senderKey) == false)
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 해당 명렁어는 개발자 권한이 필요합니다.", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                if (strContents == "")
                {
                    userDirector.resetChattingCount();
                }
                else
                {
                    CUser userInfo = userDirector.getUserInfoByName(strContents);
                    if (userInfo.UserKey == 0)
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 대화명을 찾을 수 없습니다.", ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    userDirector.resetChattingCount(userInfo.UserKey);
                }

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[SYSTEM] 대화량 리셋 완료 됐습니다.", ParseMode.Default, false, false, iMessageID);
            }
            //========================================================================================
            // 대화량 순위
            //========================================================================================
            else if (strCommend == "/순위")
            {
                var allUserInfo = userDirector.getAllUserInfo();
                Dictionary<long, ulong> dicChattingCount = new Dictionary<long, ulong>();
                ulong totalCount = 0;
                int senderRank = 0;
                ulong senderChattingCount = 0;
                string senderName = "";

                foreach (var iter in allUserInfo)
                {
                    dicChattingCount.Add(iter.Value.UserKey, iter.Value.chattingCount);
                    totalCount += iter.Value.chattingCount;

                    if (iter.Value.UserKey == senderKey)
                    {
                        CUser senderInfo = userDirector.getUserInfo(iter.Value.UserKey);
                        senderName = senderInfo.Name;
                        senderChattingCount = iter.Value.chattingCount;
                    }
                }

                int rank = 1;
                int afterRank = 1;
                ulong afterValue = 0;
                bool isContinue = false;
                strPrint += "[ 대화량 순위 Top 10 ]\n============================\n";
                foreach (KeyValuePair<long, ulong> item in dicChattingCount.OrderByDescending(key => key.Value))
                {
                    if (item.Value == 0)
                        break;

                    var user = userDirector.getUserInfo(item.Key);
                    double value = Convert.ToDouble(item.Value);
                    double count = Convert.ToDouble(totalCount);

                    if (afterValue == item.Value)
                    {
                        if (isContinue == false)
                            afterRank--;

                        if (item.Key == senderKey)
                            senderRank = afterRank;

                        isContinue = true;
                        if (rank <= 10)
                            strPrint += afterRank.ToString() + ". " + user.Name + " - " + item.Value + "건 (" + Math.Round(value / count * 100.0, 2) + "%)\n";
                        rank++;
                    }
                    else
                    {
                        if (item.Key == senderKey)
                            senderRank = rank;

                        isContinue = false;
                        if (rank <= 10)
                            strPrint += rank.ToString() + ". " + user.Name + " - " + item.Value + "건 (" + Math.Round(value / count * 100.0, 2) + "%)\n";
                        rank++;
                        afterRank = rank;
                    }

                    afterValue = item.Value;
                }

                double countDouble = Convert.ToDouble(totalCount);
                strPrint += "============================\n";
                strPrint += "* 내 순위\n" + senderRank.ToString() + ". " + senderName + " - " + senderChattingCount.ToString() + "건 (" + Math.Round(senderChattingCount / countDouble * 100.0, 2) + "%)\n";

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
            }
            //========================================================================================
            // 안내
            //========================================================================================
            else if (strCommend == "/안내")
            {
                await Bot.SendChatActionAsync(varMessage.Chat.Id, ChatAction.UploadPhoto);

                const string strCDTInfo01 = @"CDT_Info/CDT_Info_1.png";
                const string strCDTInfo02 = @"CDT_Info/CDT_Info_2.png";
                const string strCDTInfo03 = @"CDT_Info/CDT_Info_3.png";
                const string strCDTInfo04 = @"CDT_Info/CDT_Info_4.png";
                const string strCDTInfo05 = @"CDT_Info/CDT_Info_5.png";
                const string strDiscordGuide = @"CDT_Info/DiscordGuide.png";

                var fileName01 = strCDTInfo01.Split(Path.DirectorySeparatorChar).Last();
                var fileName02 = strCDTInfo02.Split(Path.DirectorySeparatorChar).Last();
                var fileName03 = strCDTInfo03.Split(Path.DirectorySeparatorChar).Last();
                var fileName04 = strCDTInfo04.Split(Path.DirectorySeparatorChar).Last();
                var fileName05 = strCDTInfo05.Split(Path.DirectorySeparatorChar).Last();
                var fileDiscordGuide = strDiscordGuide.Split(Path.DirectorySeparatorChar).Last();

                var fileStream01 = new FileStream(strCDTInfo01, FileMode.Open, FileAccess.Read, FileShare.Read);
                var fileStream02 = new FileStream(strCDTInfo02, FileMode.Open, FileAccess.Read, FileShare.Read);
                var fileStream03 = new FileStream(strCDTInfo03, FileMode.Open, FileAccess.Read, FileShare.Read);
                var fileStream04 = new FileStream(strCDTInfo04, FileMode.Open, FileAccess.Read, FileShare.Read);
                var fileStream05 = new FileStream(strCDTInfo05, FileMode.Open, FileAccess.Read, FileShare.Read);
                var fileDiscordStream = new FileStream(strDiscordGuide, FileMode.Open, FileAccess.Read, FileShare.Read);

                await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream01, "");
                await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream02, "");
                await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream03, "");
                await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream04, "");
                await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream05, "");
                await Bot.SendPhotoAsync(varMessage.Chat.Id, fileDiscordStream, "");

                strPrint = "위 가이드는 본방에서\n/안내 입력으로 다시 보실 수 있습니다.\n\n";
                foreach (var iter in config.admin_ID_)
                {
                    strPrint += "@" + iter + " ";
                }

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);

                if (DateTime.Now.Hour >= 0 && DateTime.Now.Hour < 8)
                {
                    strPrint = "현재 운영진 업무 시간이 종료되었습니다.\n가입절차 안내와 문의사항 답변은 잠시 기다려 주세요.\n\n[운영진 주 업무시간: 18시 ~자정]\n- 18시 이전 문의 가능, 단 답변이 느릴 수 있음\n- 새벽 텔레그램 1:1 문의 자제";
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
                }
            }
            //========================================================================================
            // 전달
            //========================================================================================
            else if (strCommend == "/전달")
            {
                if (config.isDeveloper(senderKey) == false)
                    return;

                await Bot.SendTextMessageAsync(config.getGroupKey(GROUP_TYPE.GROUP_TYPE_CLAN), strContents);
            }
            //========================================================================================
            // 리포트
            //========================================================================================
            else if (strCommend == "/리포트")
            {
                //await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
            }
            //========================================================================================
            // 상태
            //========================================================================================
            else if (strCommend == "/상태")
            {
                strPrint += "Running.......\n";
                strPrint += "[System Time] " + systemInfo.GetNowTime() + "\n";
                strPrint += "[Running Time] " + systemInfo.GetRunningTime() + "\n";

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
            }

            strPrint = "";
        }
    }
}
