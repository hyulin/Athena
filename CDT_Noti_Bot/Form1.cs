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
            timer.Interval = 1000 * 60 * 5; // 5분
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
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
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

        private void Form1_Load(object sender, EventArgs e)
        {
            telegramAPIAsync();

            setTelegramEvent();      //위에 만든 소스에 이어서 추가해주세요.
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
                varMessage.Text = "/안내";
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
                strPrint += "========================================\n";
                strPrint += "[ Clien Delicious Team Notice Bot v1.0 ]\n\n";
                strPrint += "/공지 : 팀 공지사항을 출력합니다.\n";
                strPrint += "/모임 : 모임 공지와 참가자를 출력합니다.\n";
                strPrint += "      - /참가|대화명 : 모임에 참가신청합니다.\n";
                strPrint += "      - /취소|대화명 : 모임에 참가신청을 취소합니다.\n";
                strPrint += "/안내 : 팀 안내 메시지를 출력합니다.\n";
                strPrint += "/리포트 : 업데이트 내역, 개발 예정 항목을 출력합니다.\n";
                strPrint += "/상태 : 현재 봇 상태를 출력합니다. 대답이 없으면 이상.\n";
                strPrint += "----------------------------------------\n";
                strPrint += "CDT 1대 운영자 : 냉각콜라, 휴린, 청포도, 만슬\n";
                strPrint += "========================================\n";
                strPrint += "버그 및 문의사항이 있으시면 '휴린'에게 문의해주세요. :)\n";
                strPrint += "========================================\n";

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
            }
            else if (strOutput[0] == "/운영자도움말")
            {
                strPrint += "========================================\n";
                strPrint += "[ Clien Delicious Team Notice Bot v1.0 입니다. ]\n\n";
                strPrint += "/공지 : 팀 공지사항을 출력합니다.\n";
                strPrint += "      - /공지등록|내용 : 팀 공지사항을 등록합니다.\n";
                strPrint += "      - /공지삭제 : 팀 공지사항을 삭제합니다.\n";
                strPrint += "/모임 : 모임 공지와 참가자를 출력합니다.\n";
                strPrint += "      - /모임등록|내용 : 모임 공지를 등록합니다.\n";
                strPrint += "      - /삭제 : 모임 공지를 삭제합니다.\n";
                strPrint += "      - /참가|대화명 : 모임에 참가신청합니다.\n";
                strPrint += "      - /취소|대화명 : 모임에 참가신청을 취소합니다.\n";
                //strPrint += "      - /강제참가 : 모임 참가자를 입력합니다.\n";
                //strPrint += "      - /강제취소 : 모임 참가자를 강제로 참가취소합니다.\n";
                strPrint += "/안내 : 팀 안내 메시지를 출력합니다.\n";
                strPrint += "/리포트 : 봇 업데이트 내역, 개발 예정인 항목\n";
                strPrint += "/상태 : 현재 봇 상태를 출력합니다. 대답이 없으면 이상.\n";
                strPrint += "----------------------------------------\n";
                strPrint += "버그 및 문의사항이 있으시면 '휴린'에게 문의해주세요. :)\n";
                strPrint += "========================================\n";

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
                IList<IList<Object>> values = response.Values;
                if (values != null && values.Count > 0)
                {
                    strPrint += "#공지사항\n\n";

                    foreach (var row in values)
                    {
                        strPrint += "* " + row[0] + "\n\n";
                    }
                }

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
            }
            //else if (strOutput[0] == "/공지등록")
            //{
            //    strMassage = strMassage.Replace(strOutput[0], "");

            //    if (strOutput[1] == "")
            //    {
            //        strPrint += "[ERROR] 공지 내용이 없습니다.";
            //    }
            //    else
            //    {
            //        System.IO.File.WriteAllText(@"_Notice.txt", strOutput[1], Encoding.Unicode);
            //        strPrint += "[SUCCESS] 공지 등록완료.";
            //    }

            //    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
            //}
            //else if (strOutput[0] == "/공지삭제")
            //{
            //    strMassage = strMassage.Replace(strOutput[0], "");
            //    string strNoticeValue = System.IO.File.ReadAllText(@"_Notice.txt");

            //    if (strNoticeValue == "")
            //    {
            //        strPrint += "[ERROR] 현재 공지가 등록되지 않았습니다.";
            //    }
            //    else
            //    {
            //        System.IO.File.WriteAllText(@"_Notice.txt", "", Encoding.Unicode);
            //        strPrint += "[SUCCESS] 공지 삭제완료.";
            //    }

            //    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
            //}
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
                    System.IO.StreamReader file = new System.IO.StreamReader(@"_MeetingMember.txt");
                    string line = "";

                    strPrint += strMeetingValue;
                    strPrint += "\n----------------------------------------\n";
                    strPrint += "★ 참가자 : ";
                    strPrint += file.ReadLine();

                    while ((line = file.ReadLine()) != null)
                    {
                        strPrint += " , " + line;
                    }

                    file.Close();

                    strPrint += "\n----------------------------------------";
                }

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
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
            //else if (strOutput[0] == "/참가")
            //{
            //    strMassage = strMassage.Replace(strOutput[0], "");
            //    string strMeetingValue = System.IO.File.ReadAllText(@"_Meeting.txt");

            //    if (strMeetingValue == "")
            //    {
            //        strPrint += "[ERROR] 현재 모임이 등록되지 않았습니다.";
            //    }
            //    else
            //    {
            //        System.IO.StreamReader file = new System.IO.StreamReader(@"_MeetingMember.txt");
            //        string line;
            //        bool IsReduplication = false;

            //        while ((line = file.ReadLine()) != null)
            //        {
            //            System.Console.WriteLine(line);

            //            if (line == strUserName)
            //            {
            //                IsReduplication = true;
            //                break;
            //            }
            //        }

            //        file.Close();

            //        if (IsReduplication)
            //        {
            //            strPrint += "[ERROR] 이미 신청되어있습니다. : " + strUserName;
            //        }
            //        else
            //        {
            //            string strNoticeValue = System.IO.File.ReadAllText(@"_MeetingMember.txt");
            //            System.IO.File.AppendAllText(@"_MeetingMember.txt", strUserName + "\n", Encoding.Unicode);

            //            strPrint += "[SUCCESS] 모임 신청완료. : " + strUserName;
            //        }
            //    }

            //    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
            //}
            else if (strOutput[0] == "/참가")
            {
                strMassage = strMassage.Replace(strOutput[0], "");
                string strMeetingValue = System.IO.File.ReadAllText(@"_Meeting.txt");
                strUserName = strOutput[1];

                if (strMeetingValue == "")
                {
                    strPrint += "[ERROR] 현재 모임이 등록되지 않았습니다.";
                }
                else
                {
                    System.IO.StreamReader file = new System.IO.StreamReader(@"_MeetingMember.txt");
                    string line;
                    bool IsReduplication = false;

                    while ((line = file.ReadLine()) != null)
                    {
                        System.Console.WriteLine(line);

                        if (line == strUserName)
                        {
                            IsReduplication = true;
                            break;
                        }
                    }

                    file.Close();

                    if (IsReduplication)
                    {
                        strPrint += "[ERROR] 이미 신청되어있습니다. (" + strUserName + ")";
                    }
                    else
                    {
                        string strNoticeValue = System.IO.File.ReadAllText(@"_MeetingMember.txt");
                        System.IO.File.AppendAllText(@"_MeetingMember.txt", strUserName + "\n", Encoding.Unicode);

                        strPrint += "[SUCCESS] 모임 신청완료. (" + strUserName + ")";
                    }
                }

                await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
            }
            //else if (strOutput[0] == "/취소")
            //{
            //    strMassage = strMassage.Replace(strOutput[0], "");
            //    string strMeetingValue = System.IO.File.ReadAllText(@"_Meeting.txt");

            //    if (strMeetingValue == "")
            //    {
            //        strPrint += "[ERROR] 현재 모임이 등록되지 않았습니다.";
            //    }
            //    else
            //    {
            //        strMassage = strMassage.Replace(strOutput[0], "");
            //        System.IO.StreamReader file = new System.IO.StreamReader(@"_MeetingMember.txt");
            //        string line;
            //        string strNewMember = "";
            //        bool IsAttend = true;

            //        while ((line = file.ReadLine()) != null)
            //        {
            //            System.Console.WriteLine(line);

            //            if (line != strUserName)
            //            {
            //                strNewMember += line + "\n";
            //            }
            //            else
            //            {
            //                IsAttend = false;
            //            }
            //        }

            //        file.Close();

            //        if (IsAttend)
            //        {
            //            strPrint += "[ERROR] 모임에 신청하지 않았습니다. : " + strUserName;
            //        }
            //        else
            //        {
            //            System.IO.File.WriteAllText(@"_MeetingMember.txt", strNewMember, Encoding.Unicode);
            //            strPrint += "[SUCCESS] 모임 취소완료. : " + strUserName;
            //        }
            //    }

            //    await Bot.SendTextMessageAsync(varMessage.Chat.Id, strPrint);
            //}
            else if (strOutput[0] == "/취소")
            {
                strMassage = strMassage.Replace(strOutput[0], "");
                string strMeetingValue = System.IO.File.ReadAllText(@"_Meeting.txt");
                strUserName = strOutput[1];

                if (strMeetingValue == "")
                {
                    strPrint += "[ERROR] 현재 모임이 등록되지 않았습니다.";
                }
                else
                {
                    strMassage = strMassage.Replace(strOutput[0], "");
                    System.IO.StreamReader file = new System.IO.StreamReader(@"_MeetingMember.txt");
                    string line;
                    string strNewMember = "";
                    bool IsAttend = true;

                    while ((line = file.ReadLine()) != null)
                    {
                        System.Console.WriteLine(line);

                        if (line != strUserName)
                        {
                            strNewMember += line + "\n";
                        }
                        else
                        {
                            IsAttend = false;
                        }
                    }

                    file.Close();

                    if (IsAttend)
                    {
                        strPrint += "[ERROR] 모임에 신청하지 않았습니다. (" + strUserName + ")";
                    }
                    else
                    {
                        System.IO.File.WriteAllText(@"_MeetingMember.txt", strNewMember, Encoding.Unicode);
                        strPrint += "[SUCCESS] 모임 취소완료. (" + strUserName + ")";
                    }
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
