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

namespace CDT_Noti_Bot
{
    public partial class Form1 : Form
    {
        CSystemInfo systemInfo = new CSystemInfo();     // 시스템 정보

        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Clien Delicious Team Notice Bot";
        UserCredential credential;
        SheetsService service;
        CNotice Notice = new CNotice();
        CNotice NewNotice = new CNotice();
        CEasterEgg EasterEgg = new CEasterEgg();
        bool bRun = false;

        // Bot Token
#if DEBUG
        const string strBotToken = "624245556:AAHJQ3bwdUB6IRf1KhQ2eAg4UDWB5RTiXzI";     // 테스트 봇 토큰
#else
        const string strBotToken = "648012085:AAHxJwmDWlznWTFMNQ92hJyVwsB_ggJ9ED8";     // 봇 토큰
#endif

        private Telegram.Bot.TelegramBotClient Bot = new Telegram.Bot.TelegramBotClient(strBotToken);

        public Form1()
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

            // 타이머 생성 및 시작
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 5000; // 5초
            timer.Elapsed += new ElapsedEventHandler(timer_Elapsed);
            timer.Start();

            InitializeComponent();
        }

        // 쓰레드풀의 작업쓰레드가 지정된 시간 간격으로
        // 아래 이벤트 핸들러 실행
        public void timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            string strPrint = "";

            // Define request parameters.
            String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
            String range = "클랜 공지!C15:C23";
            String updateRange = "클랜 공지!H14";
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
            SpreadsheetsResource.ValuesResource.GetRequest updateRequest = service.Spreadsheets.Values.Get(spreadsheetId, updateRange);

            ValueRange response = request.Execute();

            if (bRun == true)
            {
                ValueRange updateResponse = updateRequest.Execute();
                bRun = false;

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

                        NewNotice.SetNotice(strPrint);

                        if (Notice.GetNotice() != NewNotice.GetNotice())
                        {
                            Notice.SetNotice(NewNotice.GetNotice());

                            Bot.SendTextMessageAsync(-1001312491933, strPrint);  // 운영진방
                            Bot.SendTextMessageAsync(-1001202203239, strPrint);  // 클랜방
                        }
                    }
                }
            }
            else
            {
                bRun = true;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            telegramAPIAsync();

            setTelegramEvent();
        }

        private void setTelegramEvent()
        {
            Bot.OnMessage += Bot_OnMessage;     // 이벤트를 추가해줍니다. 

            Bot.StartReceiving();               // 이 함수가 실행이 되어야 사용자로부터 메세지를 받을 수 있습니다.
        }

        //Events...
        // Telegram...
        private async void Bot_OnMessage(object sender, Telegram.Bot.Args.MessageEventArgs e)
        {
            var varMessage = e.Message;
            
            if (varMessage == null || (varMessage.Type != MessageType.Text && varMessage.Type != MessageType.ChatMembersAdded))
            {
                return;
            }

            string strFirstName = varMessage.From.FirstName;
            string strLastName = varMessage.From.LastName;
            int iMessageID = varMessage.MessageId;

            if (varMessage.ReplyToMessage != null && varMessage.ReplyToMessage.From.FirstName.Contains("아테나") == true)
            {
                await Bot.SendTextMessageAsync(varMessage.Chat.Id, EasterEgg.GetEasterEgg(), ParseMode.Default, false, false, iMessageID);
                return;
            }

            // 입장 메시지 일 경우
            if (varMessage.Type == MessageType.ChatMembersAdded)
            {
                if (varMessage.Chat.Title == "CDT 사전 활동안내")
                {
                    varMessage.Text = "/안내";
                }
                else if (varMessage.Chat.Title == "클리앙 딜리셔스 팀 (CDT)")
                {
                    string strInfo = "";

                    strInfo += "\n안녕하세요.\n";
                    strInfo += "서로의 삶에 힘이 되는 오버워치 클랜,\n";
                    strInfo += "'클리앙 딜리셔스 팀'에 오신 것을 환영합니다.\n\n";
                    strInfo += "저는 팀의 운영 봇인 아테나입니다.\n";
                    strInfo += "\n";
                    strInfo += "클랜 생활에 불편하신 점이 있으시거나\n";
                    strInfo += "건의사항, 문의사항이 있으실 때는\n";
                    strInfo += "운영자 냉각콜라(@Seungman),\n";
                    strInfo += "운영자 만슬(@mans3ul)에게 문의해주세요.\n";
                    strInfo += "\n";
                    strInfo += "우리 클랜의 모든 일정관리 및 운영은\n";
                    strInfo += "통합문서를 통해 확인 하실 수 있습니다.\n";
                    strInfo += "(https://goo.gl/nurbLT [딜리셔스.kr])\n";
                    strInfo += "통합 문서에 대해 문의사항이 있으실 때는\n";
                    strInfo += "운영자 청포도(@leetk321)에게 문의해주세요.\n";
                    strInfo += "\n";
                    strInfo += "클랜원들의 편의를 위한\n";
                    strInfo += "저, 아테나의 기능을 확인하시려면\n";
                    strInfo += "/도움말 을 입력해주세요.\n";
                    strInfo += "편리한 기능들이 많이 있으며,\n";
                    strInfo += "앞으로 더 추가될 예정입니다.\n";
                    strInfo += "아테나에 대해 문의사항이 있으실 때는\n";
                    strInfo += "운영자 휴린(@hyulin)에게 문의해주세요.\n";
                    strInfo += "\n";
                    strInfo += "저희 CDT에서 즐거운 오버워치 생활,\n";
                    strInfo += "그리고 더 나아가 즐거운 라이프를\n";
                    strInfo += "즐기셨으면 합니다.\n\n";
                    strInfo += "잘 부탁드립니다 :)\n";

                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strInfo);

                    return;
                }
                else
                {
                    return;
                }
            }

            //if ( (varMessage.Chat.Id == -1001312491933) || (varMessage.Chat.Id == -1001202203239) )
            //{
            //    if (varMessage.Text.Contains("#공지사항"))
            //    {
            //        await Bot.PinChatMessageAsync(varMessage.Chat.Id, varMessage.MessageId);
            //    }
            //}

            string strMassage = varMessage.Text;
            string strUserName = varMessage.Chat.FirstName + varMessage.Chat.LastName;
            string strCommend = "";
            string strContents = "";

            if (strMassage.Substring(0, 1) != "/")
            {
                return;
            }

            if (strMassage.IndexOf(" ") == -1)
            {
                strCommend = strMassage;
            }
            else
            {
                strCommend = strMassage.Substring(0, strMassage.IndexOf(" "));
                strContents = strMassage.Substring(strMassage.IndexOf(" ") + 1, strMassage.Count() - strMassage.IndexOf(" ") - 1);
            }
            
            string strPrint = "";

            //========================================================================================
            // 공지사항 관련 명령어
            //========================================================================================
            if (strCommend == "/도움말")
            {
                strPrint += "==================================\n";
                strPrint += "[ 아테나 v1.2 ]\n[ Clien Delicious Team Notice Bot ]\n\n";
                strPrint += "/공지 : 팀 공지사항을 출력합니다.\n";
                strPrint += "/조회 검색어 : 클랜원을 조회합니다.\n";
                strPrint += "               (검색범위 : 대화명, 배틀태그, 부계정)\n";
                strPrint += "/영상 : 영상이 있던 날짜를 조회합니다.\n";
                strPrint += "       - /영상 날짜 : 플레이 영상을 조회합니다. (/영상 181006)\n";
                strPrint += "/검색 검색어 : 포지션, 모스트별로 클랜원을 검색합니다.\n";
                strPrint += "/모임 : 모임 공지와 참가자를 출력합니다.\n";
                strPrint += "/안내 : 팀 안내 메시지를 출력합니다.\n";
                strPrint += "/리포트 : 업데이트 내역, 개발 예정 항목을 출력합니다.\n";
                strPrint += "/상태 : 현재 봇 상태를 출력합니다. 대답이 없으면 이상.\n";
                strPrint += "----------------------------------\n";
                strPrint += "CDT 1대 운영자 : 냉각콜라, 휴린, 청포도, 만슬\n";
                strPrint += "==================================\n";
                strPrint += "버그 및 문의사항이 있으시면 '휴린'에게 문의해주세요. :)\n";
                strPrint += "==================================\n";

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
            }
            else if (strCommend == "/운영자도움말")
            {
                strPrint += "==================================\n";
                strPrint += "[ 아테나 v1.2 ]\n[ Clien Delicious Team Notice Bot ]\n\n";
                strPrint += "/공지 : 팀 공지사항을 출력합니다.\n";
                strPrint += "/조회 검색어 : 클랜원을 조회합니다. (검색범위 : 대화명, 배틀태그)\n";
                strPrint += "/영상 : 영상이 있던 날짜를 조회합니다.\n";
                strPrint += "       - /영상 날짜 : 플레이 영상을 조회합니다. (/영상 181006)\n";
                strPrint += "/검색 검색어 : 포지션, 모스트별로 클랜원을 검색합니다.\n";
                strPrint += "/모임 : 모임 공지와 참가자를 출력합니다.\n";
                strPrint += "       - /모임등록 내용 : 모임 공지를 등록합니다.\n";
                strPrint += "       - /모임삭제 : 모임 공지를 삭제합니다.\n";
                strPrint += "/안내 : 팀 안내 메시지를 출력합니다.\n";
                strPrint += "/리포트 : 봇 업데이트 내역, 개발 예정인 항목\n";
                strPrint += "/상태 : 현재 봇 상태를 출력합니다. 대답이 없으면 이상.\n";
                strPrint += "----------------------------------\n";
                strPrint += "버그 및 문의사항이 있으시면 '휴린'에게 문의해주세요. :)\n";
                strPrint += "==================================\n";

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
            }
            //========================================================================================
            // 공지사항 관련 명령어
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

                System.IO.File.WriteAllText(@"_Notice.txt", strPrint, Encoding.Unicode);

                if (strPrint != "")
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 공지가 등록되지 않았습니다.", ParseMode.Default, false, false, iMessageID);
                }
            }
            //========================================================================================
            // 조회 관련 명령어
            //========================================================================================
            else if (strCommend == "/조회")
            {
                if (strContents == "")
                {
                    strPrint += "[ERROR] 대화명이 없습니다.";
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                }
                else
                {
                    // Define request parameters.
                    String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
                    String range = "클랜원 목록!C7:M";
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
                                if (row[0].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[1].ToString().ToUpper().Contains(strContents.ToUpper()) ||
                                    row[2].ToString().ToUpper().Contains(strContents.ToUpper()))
                                {
                                    string[] strBattleTag = row[1].ToString().Split('#');
                                    string strUrl = "http://playoverwatch.com/ko-kr/career/pc/" + strBattleTag[0] + "-" + strBattleTag[1];
                                    string strScore = "전적을 조회할 수 없습니다.";
                                    string strTier = "전적을 조회할 수 없습니다.";

                                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "'" + strBattleTag[0] + "#" + strBattleTag[1] + "'의 전적을 조회 중입니다.\n잠시만 기다려주세요.");

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
                                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "'" + strBattleTag[0] + "#" + strBattleTag[1] + "'의 전적을 조회할 수 없습니다.");
                                    }

                                    if (bContinue == true)
                                    {
                                        strPrint += "==================================\n";
                                    }

                                    strPrint += "[ " + row[0].ToString() + " ]\n";
                                    strPrint += "* 티어 및 점수는 전적을 조회합니다. *\n\n";
                                    strPrint += "1. 배틀태그 : " + row[1] + "\n";
                                    strPrint += "2. 티어 : " + strTier + "\n";
                                    strPrint += "3. 점수 : " + strScore + "\n";
                                    strPrint += "4. 부계정 배틀태그 : " + row[2] + "\n";
                                    strPrint += "5. 포지션 : " + row[3] + "\n";
                                    strPrint += "6. 모스트 : " + row[4].ToString() + " / " + row[5].ToString() + " / " + row[6].ToString() + "\n";
                                    strPrint += "7. 이외 가능 픽 : " + row[7] + "\n";
                                    strPrint += "8. 접속 시간대 : " + row[8] + "\n";
                                    strPrint += "9. 소개\n";
                                    strPrint += "\t- " + row[9] + "\n";

                                    bContinue = true;   // 한 명만 출력된다면 이 부분은 무시됨.
                                }
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
            else if (strCommend == "/영상")
            {
                // Define request parameters.
                String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
                String range = "경기 URL (18/4분기)!B5:G";
                SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                if (strContents == "")
                {
                    ValueRange response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            foreach (var row in values)
                            {
                                if (row.Count() == 6 && row[0].ToString() != "")
                                {
                                    strPrint += "[" + row[0].ToString() + "] " + row[1].ToString() + "\n";
                                }
                            }
                        }
                    }

                    strPrint += "\n/영상 날짜로 영상 주소를 조회하실 수 있습니다.\n";
                    strPrint += "(ex: /영상 181006)";
                }
                else
                {
                    string year = "20" + strContents.Substring(0, 2);
                    string month = strContents.Substring(2, 2);
                    string day = strContents.Substring(4, 2);
                    string date = year + "." + month + "." + day;
                    bool bContinue = false;
                    string user = "";

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
                                                strPrint += user + "(" + row[4].ToString() + ")" + " : " + row[5].ToString() + "\n";
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
                                                strPrint += row[3].ToString() + "(" + row[4].ToString() + ")" + " : " + row[5].ToString() + "\n";
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

                if (strPrint != "")
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 영상을 찾을 수 없습니다.", ParseMode.Default, false, false, iMessageID);
                }
            }
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
                    String range = "클랜원 목록!C7:M";
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
            // 정모 관련 명령어
            //========================================================================================
            else if (strCommend == "/모임")
            {
                string strMeetingValue = System.IO.File.ReadAllText(@"_Meeting.txt");

                if (strMeetingValue == "")
                {
                    strPrint += "[ERROR] 현재 모임이 등록되지 않았습니다.";
                }
                else
                {
                    strPrint += strMeetingValue + "\n";
                    strPrint += "\n----------------------------------------\n";
                    strPrint += "★ 참가자 : ";

                    // Define request parameters.
                    String spreadsheetId = "17G2eOb0WH5P__qFOthhqJ487ShjCtvJ6GpiUZ_mr5B8";
                    String range = "10월 정모추진 (모집중)!C12:C";
                    SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                    ValueRange response = request.Execute();
                    if (response != null)
                    {
                        IList<IList<Object>> values = response.Values;
                        if (values != null && values.Count > 0)
                        {
                            foreach (var row in values)
                            {
                                strPrint += row[0] + " , ";
                            }
                        }

                        strPrint += "\n----------------------------------------";
                    }
                }

                if (strPrint != "")
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 모임이 등록되지 않았습니다.", ParseMode.Default, false, false, iMessageID);
                }
            }
            else if (strCommend == "/모임등록")
            {
                if (strContents == "")
                {
                    strPrint += "[ERROR] 모임 내용이 없습니다.";
                }
                else
                {
                    System.IO.File.WriteAllText(@"_Meeting.txt", strContents, Encoding.Unicode);
                    strPrint += "[SUCCESS] 모임 등록완료.";
                }

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
            }
            else if (strCommend == "/모임삭제")
            {
                string strMeetingValue = System.IO.File.ReadAllText(@"_Meeting.txt");

                if (strMeetingValue == "")
                {
                    strPrint += "[ERROR] 현재 모임이 등록되지 않았습니다.";
                }
                else
                {
                    System.IO.File.WriteAllText(@"_Meeting.txt", "", Encoding.Unicode);
                    strPrint += "[SUCCESS] 모임 삭제완료.";
                }

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
            }
            //========================================================================================
            // 안내 관련 명령어
            //========================================================================================
            else if (strCommend == "/안내")
            {
                await Bot.SendChatActionAsync(varMessage.Chat.Id, ChatAction.UploadPhoto);

                const string strCDTInfo01 = @"CDT_Info/01.jpg";
                const string strCDTInfo02 = @"CDT_Info/02.jpg";
                const string strCDTInfo03 = @"CDT_Info/03.jpg";
                const string strCDTInfo04 = @"CDT_Info/04.jpg";

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
            }
            else if (strCommend == "/리포트")
            {
                string strReportValue = System.IO.File.ReadAllText(@"_Report.txt");

                if (strReportValue == "")
                {
                    strPrint += "[ERROR] 현재 업데이트 리포트가 등록되지 않았습니다.";
                }
                else
                {
                    strPrint += strReportValue;
                }

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
            }
            else if (strCommend == "/상태")
            {
                strPrint += "Running.......\n";
                strPrint += "[System Time] " + systemInfo.GetNowTime() + "\n";
                strPrint += "[Running Time] " + systemInfo.GetRunningTime() + "\n";
                //strPrint += "[Message Request Count] " + systemInfo.GetMessageReqCount() + "\n";
                //strPrint += "[Google Sheetp API Request Count] " + systemInfo.GetGoogleSheetReqCount() + "\n";

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint, ParseMode.Default, false, false, iMessageID);
            }

            strPrint = "";
        }

        // init methods...
        private async void telegramAPIAsync()
        {
            //Bot 에 대한 정보를 가져온다.
            var me = await Bot.GetMeAsync();
        }
    }
}
