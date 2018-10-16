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

namespace CDT_Noti_Bot
{
    public partial class Form1 : Form
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Clien Delicious Team Bot";
        UserCredential credential;
        SheetsService service;
        CNotice Notice = new CNotice();
        CNotice NewNotice = new CNotice();
        bool bRun = false;

        // Bot Token
        const string strBotToken = "648012085:AAHxJwmDWlznWTFMNQ92hJyVwsB_ggJ9ED8";
        private Telegram.Bot.TelegramBotClient Bot = new Telegram.Bot.TelegramBotClient(strBotToken);

        public Form1()
        {
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

            // 입장 메시지 일 경우
            if (varMessage.Type == MessageType.ChatMembersAdded)
            {
                if (varMessage.Chat.Title == "CDT 사전 활동안내")
                {
                    varMessage.Text = "/안내";
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
            string[] strOutput = strMassage.Split('|');
            string strPrint = "";

            //========================================================================================
            // 공지사항 관련 명령어
            //========================================================================================
            if (strOutput[0] == "/도움말")
            {
                strPrint += "==================================\n";
                strPrint += "[ 아테나 v1.1 ]\n[ (Clien Delicious Team Notice Bot) ]\n\n";
                strPrint += "/공지 : 팀 공지사항을 출력합니다.\n";
                strPrint += "/조회|검색어 : 클랜원을 조회합니다.\n";
                strPrint += "               (검색범위 : 대화명, 배틀태그, 부계정)\n";
                strPrint += "/모임 : 모임 공지와 참가자를 출력합니다.\n";
                strPrint += "/안내 : 팀 안내 메시지를 출력합니다.\n";
                strPrint += "/리포트 : 업데이트 내역, 개발 예정 항목을 출력합니다.\n";
                strPrint += "/상태 : 현재 봇 상태를 출력합니다. 대답이 없으면 이상.\n";
                strPrint += "----------------------------------\n";
                strPrint += "CDT 1대 운영자 : 냉각콜라, 휴린, 청포도, 만슬\n";
                strPrint += "==================================\n";
                strPrint += "버그 및 문의사항이 있으시면 '휴린'에게 문의해주세요. :)\n";
                strPrint += "==================================\n";

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
            }
            else if (strOutput[0] == "/운영자도움말")
            {
                strPrint += "==================================\n";
                strPrint += "[ 아테나 v1.1 ]\n[ (Clien Delicious Team Notice Bot) ]\n\n";
                strPrint += "/공지 : 팀 공지사항을 출력합니다.\n";
                strPrint += "/조회|검색어 : 클랜원을 조회합니다. (검색범위 : 대화명, 배틀태그)\n";
                strPrint += "/모임 : 모임 공지와 참가자를 출력합니다.\n";
                strPrint += "      - /모임등록|내용 : 모임 공지를 등록합니다.\n";
                strPrint += "      - /모임삭제 : 모임 공지를 삭제합니다.\n";
                strPrint += "/안내 : 팀 안내 메시지를 출력합니다.\n";
                strPrint += "/리포트 : 봇 업데이트 내역, 개발 예정인 항목\n";
                strPrint += "/상태 : 현재 봇 상태를 출력합니다. 대답이 없으면 이상.\n";
                strPrint += "----------------------------------\n";
                strPrint += "버그 및 문의사항이 있으시면 '휴린'에게 문의해주세요. :)\n";
                strPrint += "==================================\n";

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
            }
            //========================================================================================
            // 공지사항 관련 명령어
            //========================================================================================
            else if (strOutput[0] == "/공지")
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
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 공지가 등록되지 않았습니다.");
                }
            }
            //========================================================================================
            // 조회 관련 명령어
            //========================================================================================
            else if (strOutput[0] == "/조회")
            {
                if (strOutput[1] == "")
                {
                    strPrint += "[ERROR] 대화명이 없습니다.";
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
                            foreach (var row in values)
                            {
                                if (row[0].ToString().Contains(strOutput[1]) || row[1].ToString().Contains(strOutput[1]) || row[2].ToString().Contains(strOutput[1]))
                                {
                                    strPrint += "==================================\n";
                                    strPrint += "1. 클랜방 대화명 : " + row[0] + "\n";
                                    strPrint += "2. 배틀태그 : " + row[1] + "\n";
                                    strPrint += "3. 부계정 배틀태그 : " + row[2] + "\n";
                                    strPrint += "4. 포지션 : " + row[3] + "\n";
                                    strPrint += "5. 모스트 : " + row[4].ToString() + " / " + row[5].ToString() + " / " + row[6].ToString() + "\n";
                                    strPrint += "6. 이외 가능 픽 : " + row[7] + "\n";
                                    strPrint += "7. 접속 시간대 : " + row[8] + "\n";
                                    strPrint += "8. 소개\n";
                                    strPrint += "\t- " + row[9] + "\n";
                                }
                            }
                        }
                    }

                    if (strPrint != "")
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
                    }
                    else
                    {
                        await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 클랜원을 찾을 수 없습니다.");
                    }
                }
            }
            //========================================================================================
            // 정모 관련 명령어
            //========================================================================================
            else if (strOutput[0] == "/모임")
            {
                strMassage = strMassage.Replace(strOutput[0], "");
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
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
                }
                else
                {
                    await Bot.SendTextMessageAsync(varMessage.Chat.Id, "[ERROR] 모임이 등록되지 않았습니다.");
                }
            }
            else if (strOutput[0] == "/모임등록")
            {
                strMassage = strMassage.Replace(strOutput[0], "");

                if (strOutput[1] == "")
                {
                    strPrint += "[ERROR] 모임 내용이 없습니다.";
                }
                else
                {
                    System.IO.File.WriteAllText(@"_Meeting.txt", strOutput[1], Encoding.Unicode);
                    strPrint += "[SUCCESS] 모임 등록완료.";
                }

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
            }
            else if (strOutput[0] == "/모임삭제")
            {
                strMassage = strMassage.Replace(strOutput[0], "");
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

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
            }
            //========================================================================================
            // 안내 관련 명령어
            //========================================================================================
            else if (strOutput[0] == "/안내")
            {
                strMassage = strMassage.Replace(strOutput[0], "");
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
            else if (strOutput[0] == "/리포트")
            {
                strMassage = strMassage.Replace(strOutput[0], "");
                string strReportValue = System.IO.File.ReadAllText(@"_Report.txt");

                if (strReportValue == "")
                {
                    strPrint += "[ERROR] 현재 업데이트 리포트가 등록되지 않았습니다.";
                }
                else
                {
                    strPrint += strReportValue;
                }

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
            }
            else if (strOutput[0] == "/상태")
            {
                strMassage = strMassage.Replace(strOutput[0], "");

                strPrint += "정상";

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
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
