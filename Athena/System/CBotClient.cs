using System;
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

        bool isGoodMorning = false;

        // Bot Token
#if DEBUG
        const string strBotToken = "624245556:AAHJQ3bwdUB6IRf1KhQ2eAg4UDWB5RTiXzI";     // 테스트 봇 토큰
#else
        const string strBotToken = "648012085:AAHxJwmDWlznWTFMNQ92hJyVwsB_ggJ9ED8";     // 봇 토큰
#endif

        private Telegram.Bot.TelegramBotClient Bot = new Telegram.Bot.TelegramBotClient(strBotToken);

        public void InitBotClient()
        {
            systemInfo.SetStartTime();

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


            // 시트에서 유저 정보를 Load
            loadUserInfo();


            // 백업 파일에서 알림, 메모를 Load
            loadData();


            // NAS 경로 기본값 설정
            nasInfo.CurrentPath = @"D:\CDT\";


            // 아테나 구동 알림
            string strPrint = "";
            strPrint += "Ahtena Start.\n";
            strPrint += "[System Time] " + systemInfo.GetNowTime() + "\n";
            strPrint += "[Running Time] " + systemInfo.GetRunningTime() + "\n";

//#if DEBUG
//            Bot.SendTextMessageAsync(-1001219697643, strPrint);  // 운영진방
//#else
//            Bot.SendTextMessageAsync(-1001202203239, strPrint);  // 클랜방
//#endif


            // 타이머 생성 및 시작
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 5000; // 5초
            timer.Elapsed += new ElapsedEventHandler(timer_Elapsed);
            timer.Start();
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

        public void loadUserInfo()
        {
            // Define request parameters.
            String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
            String range = "클랜원 목록!C7:N";
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
                        if (row[10].ToString() == "")
                            continue;

                        long userKey = Convert.ToInt64(row[10].ToString());
                        var userData = userDirector.getUserInfo(userKey);
                        if (userData.UserKey != 0)
                        {
                            // 이미 등록한 유저. 갱신한다.
                            isReflesh = true;
                        }

                        CUser user = new CUser();
                        user = setUserInfo(row, Convert.ToInt64(row[10].ToString()));

                        // 휴린, 냉각콜라, 만슬, 루미녹스일 경우
                        if ((user.UserKey == 23842788) || (user.UserKey == 50872681) ||
                            (user.UserKey == 474057213) || (user.UserKey == 83970696))
                        {
                            // 유저 타입을 관리자로
                            user.UserType = USER_TYPE.USER_TYPE_ADMIN;
                        }

                        if (isReflesh == false)
                            userDirector.addUserInfo(userKey, user);
                        else
                            userDirector.reflechUserInfo(userKey, user);
                    }
                }
            }
        }

        public void loadData()
        {
            string FolderName = @"Data/";
            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(FolderName);
            foreach (System.IO.FileInfo File in di.GetFiles())
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
                else if (File.Name.Contains("Memo_") == true)
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

            return;
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

            if (row[3].ToString() == "플렉스")
                user.Position |= POSITION.POSITION_FLEX;
            if (row[3].ToString().ToUpper().Contains("딜"))
                user.Position |= POSITION.POSITION_DPS;
            if (row[3].ToString().ToUpper().Contains("탱"))
                user.Position |= POSITION.POSITION_TANK;
            if (row[3].ToString().ToUpper().Contains("힐"))
                user.Position |= POSITION.POSITION_SUPP;

            string[] most = new string[3];
            most[0] = row[4].ToString();
            most[1] = row[5].ToString();
            most[2] = row[6].ToString();
            user.MostPick = most;

            user.OtherPick = row[7].ToString();
            user.Time = row[8].ToString();
            user.Info = row[9].ToString();

            // 휴린, 냉각콜라, 만슬, 루미녹스일 경우
            if ((user.UserKey == 23842788) || (user.UserKey == 50872681) ||
                (user.UserKey == 474057213) || (user.UserKey == 83970696))
            {
                // 유저 타입을 관리자로
                user.UserType = USER_TYPE.USER_TYPE_ADMIN;
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

            if (DateTime.Now.Hour == 8)
            {
                if (isGoodMorning == false)
                {
                    isGoodMorning = true;
                    strPrint += "굿모닝~ 오늘도 즐거운 하루 되세요~ :)";

#if DEBUG
                    Bot.SendTextMessageAsync(-1001219697643, strPrint);  // 운영진방
#else
                    Bot.SendTextMessageAsync(-1001202203239, strPrint);  // 클랜방
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
                                Bot.SendTextMessageAsync(-1001219697643, strPrint);  // 운영진방
#else
                                Bot.SendTextMessageAsync(-1001202203239, strPrint);  // 클랜방
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
            String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
            String range = "클랜 공지!C15:C23";
            String updateRange = "클랜 공지!H14";
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
            SpreadsheetsResource.ValuesResource.GetRequest updateRequest = service.Spreadsheets.Values.Get(spreadsheetId, updateRange);

            ValueRange response = request.Execute();
            ValueRange updateResponse = updateRequest.Execute();

            if (response != null && updateResponse != null)
            {
                IList<IList<Object>> values = response.Values;
                IList<IList<Object>> updateValues = updateResponse.Values;

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

#if DEBUG
                    Bot.SendTextMessageAsync(-1001219697643, strPrint);  // 운영진방
#else
                    Bot.SendTextMessageAsync(-1001202203239, strPrint);  // 클랜방
#endif
                }
            }
        }

        //Events...
        // Telegram...
        private async void Bot_OnMessage(object sender, Telegram.Bot.Args.MessageEventArgs e)
        {
            var varMessage = e.Message;

            if (varMessage == null || (varMessage.Type != MessageType.Text && varMessage.Type != MessageType.ChatMembersAdded))
            {
                if (varMessage.Caption != "/up")
                    return;
                else
                    varMessage.Text = "/up";
            }

            DateTime convertTime = varMessage.Date.AddHours(9);
            if (convertTime < systemInfo.GetStartTimeToDate())
            {
                return;
            }

            // 입장 메시지 일 경우
            if (varMessage.Type == MessageType.ChatMembersAdded)
            {
                if (varMessage.Chat.Id == -1001389956706)   // 사전안내방
                {
                    varMessage.Text = "/안내";
                }
                else if (varMessage.Chat.Id == -1001202203239)      // 본방
                {
                    string strInfo = "";

                    strInfo += "\n안녕하세요.\n";
                    strInfo += "서로의 삶에 힘이 되는 오버워치 클랜,\n";
                    strInfo += "'클리앙 딜리셔스 팀'에 오신 것을 환영합니다.\n\n";
                    strInfo += "저는 팀의 운영 봇인 아테나입니다.\n";
                    strInfo += "\n";
                    strInfo += "클랜 생활에 불편하신 점이 있으시거나\n";
                    strInfo += "건의사항, 문의사항이 있으실 때는\n";
                    strInfo += "냉각콜라(@Seungman),\n";
                    strInfo += "휴린(@hyulin),\n";
                    strInfo += "만슬(@mans3ul),\n";
                    strInfo += "LUMINOX(@urusaikara)에게\n";
                    strInfo += "문의해주세요.\n\n";
                    strInfo += "클랜원들의 편의를 위한\n";
                    strInfo += "저, 아테나의 기능을 확인하시려면\n";
                    strInfo += "/도움말 을 입력해주세요.\n";
                    strInfo += "편리한 기능들이 많이 있으며,\n";
                    strInfo += "앞으로 더 추가될 예정입니다.\n";
                    strInfo += "아테나에 대해 문의사항이 있으실 때는\n";
                    strInfo += "운영자 휴린(@hyulin)에게 문의해주세요.\n";
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

                    return;
                }
                else
                {
                    return;
                }
            }

            // 메시지 정보 추출
            string strFirstName = varMessage.From.FirstName;
            string strLastName = varMessage.From.LastName;
            string strUserID = varMessage.From.Username;
            int iMessageID = varMessage.MessageId;
            long senderKey = varMessage.From.Id;
            DateTime time = convertTime;

            // CDT 관련방 아니면 동작하지 않도록 수정
            if (varMessage.Chat.Id != -1001202203239 &&     // 본방
                varMessage.Chat.Id != -1001219697643 &&     // 운영진방
                varMessage.Chat.Id != -1001389956706 &&     // 사전안내방
                varMessage.Chat.Username != "hyulin")
            {
                await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 사용할 수 없는 대화방입니다.", ParseMode.Default, false, false, iMessageID);
                CLog.WriteLog(varMessage.Chat.Id, senderKey, "", "[ERROR] 사용할 수 없는 대화방입니다.", "", "");
                return;
            }

            // 명령어, 서브명령어 분리
            string strMassage = varMessage.Text;
            string strUserName = varMessage.From.FirstName + varMessage.From.LastName;
            string strCommend = "";
            string strContents = "";
            bool isCommand = false;

            // 본방에 입력된 메시지를 각 유저 정보에 입력
#if DEBUG
            userDirector.addMessage(senderKey, strMassage, time);
#else
            if (senderKey != 0 && strMassage != "" && varMessage.Chat.Id == -1001202203239)
                userDirector.addMessage(senderKey, strMassage, time);
#endif

            // 명령어인지 아닌지 구분
            if (strMassage.Substring(0, 1) == "/")
            {
                isCommand = true;

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

            // 명령어가 아닐 경우와 사전안내방이 아닌 경우
            if (isCommand == false && varMessage.Chat.Id != -1001389956706)
            {
                bool isMainRoom = false;

                if (varMessage.Chat.Id == -1001202203239)
                    isMainRoom = true;

                Tuple<string, string, bool> tuple = naturalLanguage.morphemeProcessor(strMassage, userDirector.getMessage(senderKey), isMainRoom);

                // 대화
                if (tuple.Item1 != "" && tuple.Item3 == false)
                {
                    if (tuple.Item1.ToString() != "")
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, tuple.Item1.ToString(), ParseMode.Default, false, false, iMessageID);
                        CLog.WriteLog(varMessage.Chat.Id, senderKey, strUserName, strMassage, tuple.Item1.ToString(), "");
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
                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, naturalLanguage.replyCall(strMassage), ParseMode.Default, false, false, iMessageID);
                            CLog.WriteLog(varMessage.Chat.Id, senderKey, strUserName, strMassage, tuple.Item1.ToString(), "");
                        }
                    }
                }
            }
            else
            {
                CLog.WriteLog(varMessage.Chat.Id, senderKey, strUserName, strMassage, strCommend, strContents);
            }

            string strPrint = "";

            //========================================================================================
            // 도움말
            //========================================================================================
            if (strCommend == "/도움말" || strCommend == "/help" || strCommend == "/help@CDT_Noti_Bot")
            {
                strPrint += "==================================\n";
                strPrint += "[ 아테나 v2.0 ]\n[ Clien Delicious Team Notice Bot ]\n\n";
                strPrint += "/공지 : 클랜 공지사항을 출력합니다.\n";
                strPrint += "/일정 : 이번 달 클랜 일정을 확인합니다.\n";
                strPrint += "/등록 [본 계정 배틀태그] : 아테나에 등록 합니다.\n";
                strPrint += "/조회 [검색어] : 클랜원을 조회합니다.\n";
                strPrint += "               (검색범위 : 대화명, 배틀태그, 부계정)\n";
                strPrint += "/영상 [날짜] : 플레이 영상을 조회합니다. (/영상 20181006)\n";
                strPrint += "/검색 [검색어] : 포지션, 모스트별로 클랜원을 검색합니다.\n";
                strPrint += "/스크림 : 현재 모집 중인 스크림의 참가자를 출력합니다.\n";
                strPrint += "/스크림 [요일] : 현재 모집 중인 스크림에 참가신청합니다.\n";
                strPrint += "/스크림 취소 : 신청한 스크림에 참가를 취소합니다.\n";
                strPrint += "/조사 : 현재 진행 중인 일정 조사를 출력합니다.\n";
                strPrint += "/조사 [요일] : 현재 진행 중인 일정 조사에 체크합니다.\n";
                strPrint += "/모임 : 모임 공지와 참가자를 출력합니다.\n";
                strPrint += "/참가 : 모임에 참가 신청합니다.\n";
                strPrint += "/참가 확정 : 모임에 참가 확정합니다.\n";
                strPrint += "       (이미 참가일 경우 확정만 체크)\n";
                strPrint += "/불참 : 모임에 참가 신청을 취소합니다.\n";
                strPrint += "/투표 : 현재 진행 중인 투표를 출력합니다.\n";
                strPrint += "/투표 [숫자] : 현재 진행 중인 투표에 투표합니다.\n";
                strPrint += "/투표 결과 : 현재 진행 중인 투표의 결과를 출력합니다.\n";
                strPrint += "/기록 : 클랜 명예의 전당을 조회합니다.\n";
                strPrint += "/기록 [숫자] : 명예의 전당 상세내용을 조회합니다.\n";
                strPrint += "/뽑기 [항목1] [항목2] [항목3] ... : 하나를 뽑습니다.\n";
                strPrint += "/날씨 [지역] : 현재 날씨를 출력합니다.\n";
                strPrint += "/알림 [시간] [내용] : 설정한 시간에 알림을 줍니다.\n";
                strPrint += "/알림 제거 [시간] : 알림을 제거합니다.\n";
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
                    String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
                    String range = "클랜원 목록!C7:N";
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

                                    if (row[10].ToString() != "")
                                    {
                                        // 이미 값이 있으므로 갱신한다.
                                        userKey = Convert.ToInt64(row[10].ToString());
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

                                    if (row[3].ToString() == "플렉스")
                                        user.Position |= POSITION.POSITION_FLEX;
                                    if (row[3].ToString().ToUpper().Contains("딜"))
                                        user.Position |= POSITION.POSITION_DPS;
                                    if (row[3].ToString().ToUpper().Contains("탱"))
                                        user.Position |= POSITION.POSITION_TANK;
                                    if (row[3].ToString().ToUpper().Contains("힐"))
                                        user.Position |= POSITION.POSITION_SUPP;

                                    string[] most = new string[3];
                                    most[0] = row[4].ToString();
                                    most[1] = row[5].ToString();
                                    most[2] = row[6].ToString();
                                    user.MostPick = most;

                                    user.OtherPick = row[7].ToString();
                                    user.Time = row[8].ToString();
                                    user.Info = row[9].ToString();
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
                                if (userDirector.reflechUserInfo(user.UserKey, user) == true)
                                {
                                    strPrint += "[SUCCESS] 갱신 완료됐습니다.";
                                }
                            }

                            range = "클랜원 목록!M" + (7 + searchIndex);

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
                String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
                String range = "클랜 공지!C15:C23";
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

                // Define request parameters.
                String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
                String range = "클랜 공지!I13:Q27";
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
                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 일정이 등록되지 않았습니다.", ParseMode.Default, false, false, iMessageID);
                            return;
                        }

                        // 날짜
                        for (int index = 3; index < values.Count; index += 2)
                        {
                            var row = values[index];
                            var todo = values[index + 1];

                            for (int wday = 0; wday < 8/*일주일 7일 + 빈칸 1개*/; wday++)
                            {
                                CCalendar calendar = new CCalendar();
                                string weekDay = "";

                                // 중간에 한 칸이 비어있음
                                if (wday == 1)
                                    continue;

                                if (row[wday].ToString() == "")
                                    continue;

                                switch (wday)
                                {
                                    case 0:
                                        weekDay = "일";
                                        break;
                                    case 2:
                                        weekDay = "월";
                                        break;
                                    case 3:
                                        weekDay = "화";
                                        break;
                                    case 4:
                                        weekDay = "수";
                                        break;
                                    case 5:
                                        weekDay = "목";
                                        break;
                                    case 6:
                                        weekDay = "금";
                                        break;
                                    case 7:
                                        weekDay = "토";
                                        break;
                                    default:
                                        continue;
                                }

                                calendar.Day = Convert.ToInt32(row[wday].ToString());
                                calendar.Week = weekDay;
                                calendar.Todo = todo[wday].ToString();
                                calendarDirector.addCalendar(calendar);
                            }
                        }

                        strPrint += "[ " + title[0].ToString() + " ]\n============================\n";

                        for (int i = 1; i <= calendarDirector.getCalendarCount(); i++)
                        {
                            CCalendar calendar = calendarDirector.getCalendar(i);

                            if (calendar.Todo != "")
                            {
                                strPrint += "* " + calendar.Day + "일(" + calendar.Week + ") : " + calendar.Todo + "\n";
                            }
                        }
                    }
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
                    String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
                    String range = "클랜원 목록!C7:N";
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
                                strPrint += "- 티어 : " + strTier + "\n";
                                strPrint += "- 점수 : " + strScore + "\n";
                                strPrint += "- 본 계정 배틀태그 : " + mainBattleTag + "\n";
                                strPrint += "- 부 계정 배틀태그 : " + row[2].ToString() + "\n";
                                strPrint += "- 포지션 : " + row[3].ToString() + "\n";
                                strPrint += "- 모스트 : " + row[4].ToString() + " / " + row[5].ToString() + " / " + row[6].ToString() + "\n";
                                strPrint += "- 이외 가능 픽 : " + row[7].ToString() + "\n";
                                strPrint += "- 접속 시간대 : " + row[8].ToString() + "\n";
                                strPrint += "- 소개 : " + row[9].ToString() + "\n";

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
                String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";

                if (strContents.Contains("20") == false)
                {
                    int yearIdx = 2018;
                    while (yearIdx != 0)
                    {
                        String range = "경기 URL (" + yearIdx++ + ")!B5:G";
                        SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                        try
                        {
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

                                        if (row[1].ToString().Contains(strContents) == true)
                                        {
                                            strPrint += "[" + row[0].ToString() + "] " + row[1].ToString() + "\n";
                                        }
                                    }
                                }
                            }
                        }
                        catch
                        {
                            yearIdx = 0;
                        }
                    }
                }
                else
                {
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

                        String range = "경기 URL (" + year + ")!B5:G";
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

                        String range = "경기 URL (" + year + ")!B5:G";
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
                    String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
                    String range = "클랜원 목록!C7:N";
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
                                    if (row[3].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[4].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[5].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[6].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[3].ToString() == "플렉스")
                                    {
                                        strPrint += row[0] + "(" + row[1] + ") : ";
                                        strPrint += row[3] + "(" + row[4].ToString() + "/" + row[5].ToString() + "/" + row[6].ToString() + ")\n";
                                        bResult = true;
                                    }
                                }
                                else
                                {
                                    if (row[3].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[4].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[5].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[6].ToString().ToUpper().Contains(strContents.ToUpper()))
                                    {
                                        strPrint += row[0] + "(" + row[1] + ") : ";
                                        strPrint += row[3] + " (" + row[4].ToString() + "/" + row[5].ToString() + "/" + row[6].ToString() + ")\n";
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
                // Define request parameters.
                String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
                String range = "모임!C4:R12";
                SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                ValueRange response = request.Execute();
                if (response != null)
                {
                    IList<IList<Object>> values = response.Values;
                    if (values != null && values.Count > 0)
                    {
                        // 모임이름 ~ 문의
                        foreach (var row in values)
                        {
                            if (row.Count > 0)
                            {
                                if (row[0].ToString() == "모임이름" && row[1].ToString() == "")
                                {
                                    strPrint = "[SYSTEM] 현재 예정된 모임이 없습니다.";
                                    const string meeting = @"Function/Meeting.jpg";
                                    var fileName = meeting.Split(Path.DirectorySeparatorChar).Last();
                                    var fileStream = new FileStream(meeting, FileMode.Open, FileAccess.Read, FileShare.Read);
                                    await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream, strPrint, ParseMode.Default, false, iMessageID);
                                    return;
                                }

                                if (row[0].ToString() == "프로그램" && row[1].ToString() != "")
                                {
                                    strPrint += "* " + row[0].ToString() + "\n";
                                    strPrint += "          - [" + row[1].ToString() + "] " + row[2].ToString() + " / " + row[3].ToString() + " / " + row[5].ToString() + "\n";
                                }
                                else if (row[0].ToString() == "" && row[1].ToString() != "")
                                {
                                    strPrint += "          - [" + row[1].ToString() + "] " + row[2].ToString() + " / " + row[3].ToString() + " / " + row[5].ToString() + "\n";
                                }
                                else if (row[0].ToString() != "" && row[1].ToString() != "")
                                {
                                    strPrint += "* " + row[0].ToString() + " : " + row[1].ToString() + "\n";
                                }
                            }
                        }

                        // 공지, 회비
                        foreach (var row in values)
                        {
                            if (row.Count > 0)
                            {
                                if (row[6].ToString() == "공지" && row[7].ToString() != "")
                                {
                                    strPrint += "* " + row[6].ToString() + "\n";
                                    strPrint += row[7].ToString() + "\n";
                                }
                                else if (row[6].ToString() == "회비" && row[7].ToString() != "")
                                {
                                    strPrint += "* " + row[6].ToString() + "\n";
                                    strPrint += "          - [" + row[7].ToString() + "] " + row[8].ToString() + "\n";
                                }
                                else if (row[6].ToString() == "" && row[7].ToString() != "")
                                {
                                    strPrint += "          - [" + row[7].ToString() + "] " + row[8].ToString() + "\n";
                                }
                                else if (row[6].ToString() != "" && row[7].ToString() != "")
                                {
                                    strPrint += "* " + row[6].ToString() + " : " + row[7].ToString() + "\n";
                                }
                            }
                        }
                    }
                }

                List<string> lstConfirm = new List<string>();
                List<string> lstUndefine = new List<string>();

                // Define request parameters.
                range = "모임!C16:O";
                request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                response = request.Execute();
                if (response != null)
                {
                    IList<IList<Object>> values = response.Values;
                    if (values != null && values.Count > 0)
                    {
                        foreach (var row in values)
                        {
                            if (row.Count != 0)
                            {
                                if (row.Count >= 13)
                                {
                                    if (row[12].ToString().ToUpper().Contains('O'))
                                    {
                                        string strConfirm = row[0].ToString();

                                        if (strConfirm != "")
                                        {
                                            lstConfirm.Add(strConfirm);
                                        }
                                    }
                                    else if (row[12].ToString().ToUpper().Contains('X'))
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        string strUndefine = row[0].ToString();
                                        if (strUndefine != "")
                                        {
                                            lstUndefine.Add(strUndefine);
                                        }
                                    }
                                }
                                else
                                {
                                    string strUndefine = row[0].ToString();
                                    if (strUndefine != "")
                                    {
                                        lstUndefine.Add(strUndefine);
                                    }
                                }
                            }
                        }
                    }

                    strPrint += "----------------------------------------\n";
                    strPrint += "★ 참가자\n";
                    strPrint += "- 확정 : ";
                    bool bFirst = true;

                    if (lstConfirm.Count == 0)
                    {
                        strPrint += "없음";
                    }
                    else
                    {
                        foreach (string confirm in lstConfirm)
                        {
                            if (bFirst == true)
                            {
                                strPrint += confirm;
                                bFirst = false;
                            }
                            else
                            {
                                strPrint += " , " + confirm;
                            }
                        }
                    }

                    strPrint += "\n- 미정 : ";
                    bFirst = true;

                    if (lstUndefine.Count == 0)
                    {
                        strPrint += "없음";
                    }
                    else
                    {
                        foreach (string undefine in lstUndefine)
                        {
                            if (bFirst == true)
                            {
                                strPrint += undefine;
                                bFirst = false;
                            }
                            else
                            {
                                strPrint += " , " + undefine;
                            }
                        }
                    }

                    strPrint += "\n----------------------------------------\n";
                    strPrint += "- 확정 : " + lstConfirm.Count + "명 / 미정 : " + lstUndefine.Count + "명 / 총 " + (lstConfirm.Count + lstUndefine.Count) + "명\n";
                    strPrint += "----------------------------------------";
                }

                if (strPrint != "")
                {
                    const string meeting = @"Function/Meeting.jpg";
                    var fileName = meeting.Split(Path.DirectorySeparatorChar).Last();
                    var fileStream = new FileStream(meeting, FileMode.Open, FileAccess.Read, FileShare.Read);
                    await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream, strPrint, ParseMode.Default, false, iMessageID);
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 모임이 등록되지 않았습니다.", ParseMode.Default, false, false, iMessageID);
                }
            }
            else if (strCommend == "/참가")
            {
                String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
                String range = "모임!D4:D";
                SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                ValueRange response = request.Execute();
                if (response != null)
                {
                    IList<IList<Object>> values = response.Values;
                    if (values != null && values.Count > 0)
                    {
                        var row = values[0];
                        if (row.Count == 0)
                        {
                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 모임이 등록되지 않았습니다.", ParseMode.Default, false, false, iMessageID);
                            return;
                        }
                    }
                }

                string strNickName = strFirstName + strLastName;
                int iCellIndex = 16;
                int iTempCount = 0;
                int iRealCount = 0;
                int iBlankCell = 0;
                bool isConfirm = false;
                bool isJoin = false;

                // Define request parameters.
                spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
                range = "모임!C" + iCellIndex + ":C";
                request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                response = request.Execute();
                if (response != null)
                {
                    IList<IList<Object>> values = response.Values;
                    if (values != null && values.Count > 0)
                    {
                        foreach (var row in values)
                        {
                            if (row.Count == 0)
                            {
                                if (iBlankCell == 0)
                                {
                                    iBlankCell = iTempCount;
                                }

                                iTempCount++;

                                continue;
                            }
                            else
                            {
                                if (row[0].ToString() == strNickName)
                                {
                                    iRealCount = iTempCount;
                                    isJoin = true;

                                    if (strContents == "확정")
                                    {
                                        isConfirm = true;
                                    }
                                }

                                iTempCount++;
                            }
                        }
                    }

                    if (isJoin == true && isConfirm == false)
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 이미 모임에 참가신청을 했습니다.", ParseMode.Default, false, false, iMessageID);
                        return;
                    }

                    if (isConfirm == false)
                    {
                        if (iBlankCell == 0)
                        {
                            range = "모임!C" + (iCellIndex + iRealCount) + ":C";
                        }
                        else
                        {
                            range = "모임!C" + (iCellIndex + iBlankCell) + ":C";
                        }

                        // Define request parameters.
                        ValueRange valueRange = new ValueRange();
                        valueRange.MajorDimension = "COLUMNS"; //"ROWS";//COLUMNS 

                        var oblist = new List<object>() { strNickName };
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
                            if (strContents != "확정")
                            {
                                strPrint += "[SUCCESS] 참가 신청을 완료 했습니다.";
                            }
                        }
                    }

                    if (strContents == "확정")
                    {
                        if (iBlankCell == 0)
                        {
                            range = "모임!O" + (iCellIndex + iRealCount) + ":O";
                        }
                        else
                        {
                            range = "모임!O" + (iCellIndex + iBlankCell) + ":O";
                        }

                        // Define request parameters.
                        ValueRange valueRange = new ValueRange();
                        valueRange.MajorDimension = "COLUMNS"; //"ROWS";//COLUMNS 

                        var oblist = new List<object>() { "O" };
                        valueRange.Values = new List<IList<object>> { oblist };

                        SpreadsheetsResource.ValuesResource.UpdateRequest updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);

                        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                        UpdateValuesResponse updateResponse = updateRequest.Execute();

                        if (updateResponse == null)
                        {
                            strPrint += "\n[ERROR] 참가 확정을 할 수 없습니다.";
                        }
                        else
                        {
                            strPrint += "\n[SUCCESS] 참가 확정을 완료 했습니다.";
                        }
                    }
                }

                if (strPrint != "")
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 시트를 업데이트 할 수 없습니다.", ParseMode.Default, false, false, iMessageID);
                }
            }
            else if (strCommend == "/불참")
            {
                String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
                String range = "모임!D4:D";
                SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                ValueRange response = request.Execute();
                if (response != null)
                {
                    IList<IList<Object>> values = response.Values;
                    if (values != null && values.Count > 0)
                    {
                        var row = values[0];
                        if (row.Count == 0)
                        {
                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 모임이 등록되지 않았습니다.", ParseMode.Default, false, false, iMessageID);
                            return;
                        }
                    }
                }

                string strNickName = strFirstName + strLastName;
                int iCellIndex = 16;
                int iTempCount = 0;
                bool isJoin = false;

                // Define request parameters.
                spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
                range = "모임!C" + iCellIndex + ":C";
                request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                response = request.Execute();
                if (response != null)
                {
                    IList<IList<Object>> values = response.Values;
                    if (values != null && values.Count > 0)
                    {
                        foreach (var row in values)
                        {
                            if (row.Count != 0)
                            {
                                if (row[0].ToString() == strNickName)
                                {
                                    isJoin = true;
                                    break;
                                }
                            }

                            iTempCount++;
                        }
                    }

                    if (isJoin == true)
                    {
                        range = "모임!C" + (iCellIndex + iTempCount);

                        // Define request parameters.
                        ValueRange valueRange = new ValueRange();
                        valueRange.MajorDimension = "ROWS"; //"ROWS";//COLUMNS 

                        var oblist = new List<object>() { "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
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
                            strPrint += "[SUCCESS] 참가 신청을 취소 했습니다.";
                        }
                    }
                    else
                    {
                        strPrint += "[ERROR] 참가 신청을 하지 않았습니다.";
                    }
                }

                if (strPrint != "")
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 시트를 업데이트 할 수 없습니다.", ParseMode.Default, false, false, iMessageID);
                }
            }
            //========================================================================================
            // 투표
            //========================================================================================
            else if (strCommend == "/투표")
            {
                bool isAnonymous = false;

                // Define request parameters.
                String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
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
                    String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
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

                        String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
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

                        String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
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
            else if ((strCommend == "/스크림") || (strCommend == "/오픈디비전"))
            {
                string sheetName = "";

                if (strCommend == "/스크림")
                    sheetName = "스크림";
                else
                    sheetName = "오픈디비전";

                if (strContents == "")
                {
                    String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
                    String range = sheetName + "!B2:U17";
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
                                for (int i = index; i < 15; i++)
                                {
                                    var row = values[i];

                                    if (row.Count <= 1)
                                    {
                                        continue;
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
                                }
                            }
                        }
                    }

                    if (strPrint != "")
                    {
                        strPrint += "\n신청은 /" + sheetName + " [요일] 로 해주세요.\n(ex: /" + sheetName + " 토일)";

                        string scrim = "";

                        if (sheetName == "스크림")
                            scrim = @"Function/Scrim.png";
                        else
                            scrim = @"Function/OpenDivision.png";

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
                        String sheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
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

                    String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
                    CUser user = new CUser();

                    // 클랜원 목록에서 정보 추출
                    String range = "클랜원 목록!C7:N";
                    SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                    ValueRange response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            foreach (var row in values)
                            {
                                if (row[10].ToString() == "")
                                    continue;

                                // 유저키 일치
                                if (Convert.ToInt64(row[10].ToString()) == senderKey)
                                {
                                    user = setUserInfo(row, senderKey);
                                    break;
                                }
                            }
                        }
                    }

                    int index = 0;
                    bool isInput = false;

                    range = sheetName + "!C6:C17";
                    request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                    response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            if (values.Count >= 12)
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
                    String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
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
                                    strPrint += "- " + day[i].ToString() + " : " + count[i].ToString() + "명\n";
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

                    String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
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

                    spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
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
                if ((userDirector.getUserInfo(senderKey).UserType != USER_TYPE.USER_TYPE_ADMIN) ||
                    (varMessage.Chat.Id != -1001219697643))
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "권한이 없는 유저 또는 대화방입니다.", ParseMode.Default, false, false, iMessageID);
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
                if ((userDirector.getUserInfo(senderKey).UserType != USER_TYPE.USER_TYPE_ADMIN) ||
                    (varMessage.Chat.Id != -1001219697643))
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "권한이 없는 유저 또는 대화방입니다.", ParseMode.Default, false, false, iMessageID);
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
                if ((userDirector.getUserInfo(senderKey).UserType != USER_TYPE.USER_TYPE_ADMIN) ||
                    (varMessage.Chat.Id != -1001219697643))
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "권한이 없는 유저 또는 대화방입니다.", ParseMode.Default, false, false, iMessageID);
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
                if ((userDirector.getUserInfo(senderKey).UserType != USER_TYPE.USER_TYPE_ADMIN) ||
                    (varMessage.Chat.Id != -1001219697643))
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "권한이 없는 유저 또는 대화방입니다.", ParseMode.Default, false, false, iMessageID);
                    return;
                }

                // Download File
                var file = await Bot.GetFileAsync(varMessage.Document.FileId);
                var fileName = nasInfo.CurrentPath + varMessage.Document.FileName;

                using (var saveStream = new FileStream(fileName, FileMode.CreateNew, FileAccess.Write))
                {
                    await Bot.DownloadFileAsync(file.FilePath, saveStream);
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[SYSTEM] 파일 업로드 완료.\n" + fileName, ParseMode.Default, false, false, iMessageID);
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
                        strPrint += "[ERROR] 지역을 다시 확인해주세요.";
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
                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 날씨를 조회할 수 없습니다.", ParseMode.Default, false, false, iMessageID);
                            return;
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
                            strPrint += "- 초미세먼지 : " + pm25GradeString.ToString() + "(" + pm25Value.ToString() + ")\n\n";
                            strPrint += "* 자료제공\n(날씨) OpenWeatherMap\n(대기) 한국환경공단, 공공데이터포럼";
                        }
                        catch
                        {
                            await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 날씨를 조회할 수 없습니다.", ParseMode.Default, false, false, iMessageID);
                            return;
                        }
                    }
                }

                if (strPrint != "")
                {
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
                else if (strContents.Substring(0, 2) == "제거")
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
                    string notiTime = strContents.Substring(0, 4);
                    string notiString = strContents.Substring(5);

                    int hour = Convert.ToInt32(notiTime.Substring(0, 2));
                    int min = Convert.ToInt32(notiTime.Substring(2, 2));

                    bool isSearch = false;

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

                        userDirector.RemoveMemo(senderKey, index - 1);
                        strPrint += "[SYSTEM] 해당 메모가 제거 되었습니다.";
                    }
                    else if (contents[0] != "제거" && contents[0] != "")
                    {
                        userDirector.addMemo(senderKey, strContents);
                        strPrint += "[SYSTEM] 해당 메모가 저장 되었습니다.";
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
            // 안내
            //========================================================================================
            else if (strCommend == "/안내")
            {
                await Bot.SendChatActionAsync(varMessage.Chat.Id, ChatAction.UploadPhoto);

                const string strCDTInfo01 = @"CDT_Info/CDT_Info_1.png";
                const string strCDTInfo02 = @"CDT_Info/CDT_Info_2.png";
                const string strCDTInfo03 = @"CDT_Info/CDT_Info_3.png";
                const string strCDTInfo04 = @"CDT_Info/CDT_Info_4.png";

                var fileName01 = strCDTInfo01.Split(Path.DirectorySeparatorChar).Last();
                var fileName02 = strCDTInfo02.Split(Path.DirectorySeparatorChar).Last();
                var fileName03 = strCDTInfo03.Split(Path.DirectorySeparatorChar).Last();
                var fileName04 = strCDTInfo04.Split(Path.DirectorySeparatorChar).Last();

                var fileStream01 = new FileStream(strCDTInfo01, FileMode.Open, FileAccess.Read, FileShare.Read);
                var fileStream02 = new FileStream(strCDTInfo02, FileMode.Open, FileAccess.Read, FileShare.Read);
                var fileStream03 = new FileStream(strCDTInfo03, FileMode.Open, FileAccess.Read, FileShare.Read);
                var fileStream04 = new FileStream(strCDTInfo04, FileMode.Open, FileAccess.Read, FileShare.Read);

                await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream01, "");
                await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream02, "");
                await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream03, "");
                await Bot.SendPhotoAsync(varMessage.Chat.Id, fileStream04, "");

                strPrint = "위 가이드는 본방에서 /안내 입력 시 다시 보실 수 있습니다.";

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
            }
            //========================================================================================
            // 전달
            //========================================================================================
            else if (strCommend == "/전달")
            {
                if (senderKey != 23842788)
                    return;

                await Bot.SendTextMessageAsync(-1001202203239, strContents);
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
